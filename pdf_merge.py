import os
import threading
import platform
import tkinter as tk
from tkinter import filedialog
from pypdf import PdfWriter, PdfReader
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.lang import Builder
from kivy.core.window import Window
from kivy.config import Config
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.resources import resource_add_path
from kivy.clock import Clock
from kivy.properties import ListProperty, StringProperty

if platform.system() == 'Linux':
    resource_add_path("/usr/share/fonts/opentype/noto/")
    LabelBase.register(DEFAULT_FONT, "NotoSansCJK-Regular.ttc")
else:
    resource_add_path("C:\\Windows\\Fonts")
    LabelBase.register(DEFAULT_FONT, "meiryo.ttc")

TITLE = 'PDFMerge'
VERSION = '0.5.0'

PDF_SIZES_MM = {
    'A0': (1189, 841), 'A1': (841, 594), 'A2': (594, 420),
    'A3': (420, 297), 'A4': (297, 210)
}

def convert_mm_to_pt(mm):
    return round(mm / 25.4 * 72, 1)

PDF_SIZES_PT = {k: (convert_mm_to_pt(v[0]), convert_mm_to_pt(v[1])) for k, v in PDF_SIZES_MM.items()}
SIZE_ORDER = ['A4', 'A3', 'A2', 'A1', 'A0']

# PDFファイルの処理を行うクラス
# - PDFのサイズ判定
# - カテゴリごとのマージ処理
# - 縦向き・横向きの分類
class PdfProcessor:
    # 初期化処理
    # directory: PDFファイルが格納されているディレクトリ
    # threshold: サイズ判定の閾値（デフォルト0.9）
    def __init__(self, directory, threshold=0.9):
        self.directory = directory
        self.threshold = threshold

    # PDFファイルのサイズと向きを判定する
    # file_name: 調査対象のPDFファイル名
    def get_pdf_info(self, file_name):
        pdf_path = os.path.join(self.directory, file_name)
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            size = reader.pages[0].mediabox[2:]
        
        for category, (long, short) in PDF_SIZES_PT.items():
            if (size[0] >= long * self.threshold and size[1] >= short * self.threshold):
                return size, category, 'h'
            elif (size[0] >= short * self.threshold and size[1] >= long * self.threshold):
                return size, category, 'v'
        
        return size, 'Unknown', 'Unknown'

    # PDFをカテゴリ・向きごとに結合する
    # pdf_data          : 各PDFのサイズ・カテゴリ・向きの情報
    # hv_mode           : Trueの場合は横向き・縦向きを分けて処理
    def merge_pdfs(self, pdf_data, hv_mode=True, start_size='A4', end_size='A0'):
        # 選択された範囲のサイズリストを作成
        try:
            idx1 = SIZE_ORDER.index(start_size)
            idx2 = SIZE_ORDER.index(end_size)
            # どちらが大きくても対応できるように min/max をとる
            s_idx, e_idx = min(idx1, idx2), max(idx1, idx2)
            allowed_sizes = SIZE_ORDER[s_idx : e_idx + 1] # 結合サイズの範囲の配列を作成
            print('結合範囲:', allowed_sizes)
        except ValueError:
            allowed_sizes = SIZE_ORDER # 予期せぬエラー時は全サイズ
            print('サイズ指定が無効なため、全サイズを対象にします') 

        # ファイル名の昇順でソート
        sorted_list = sorted(pdf_data.items(), key=lambda x: x[0])

        # 結合用のオブジェクトを格納する辞書（一時的なバッファ）
        active_mergers = {}

        for file_name, meta in sorted_list:
            category = meta['category']
            orientation = meta['orientation'].upper()

            # 選択された範囲に含まれるサイズのみ処理
            if category in allowed_sizes:
                # hv_mode に応じてキーを決定
                if hv_mode:
                    # 縦横を区別する場合：キーを 'H' または 'V' に集約
                    key = f"{category}_{orientation}"
                    key = f"ALL_{orientation}" #
                else:
                    # 区別しない場合：すべてのファイルを1つのキーに集約
                    key = "ALL_SIZES" #

                # キーに対応する PdfWriter がなければ作成し、ファイルを追加
                if key not in active_mergers:
                    active_mergers[key] = PdfWriter()

                file_path = os.path.join(self.directory, file_name)
                active_mergers[key].append(file_path)

        # 書き出し処理
        for key, merger in active_mergers.items():
            if merger.pages:
                # 出力ファイル名を設定
                if key == "ALL_H":
                    output_name = f'merged_{start_size}-{end_size}_Landscape.pdf'
                elif key == "ALL_V":
                    output_name = f'merged_{start_size}-{end_size}_Portrait.pdf'
                else:
                    output_name = f'merged_{start_size}-{end_size}_Combined.pdf'
                
                output_path = os.path.join(self.directory, output_name)

                if os.path.exists(output_path):
                    os.remove(output_path)
                
                merger.write(output_path)
                print(f"成功: {output_name} を作成しました。")

# PDF結合処理のメイン関数
def merge_pdfs_main(directory, hv_mode, start_size, end_size, event):
    if not os.path.isdir(directory):
        print(f"Error: Directory not found: {directory}")
        event.set() # スレッドを安全に終了させる
        return

    # 既存の結合ファイルを削除
    # 既存の結合済みPDFを削除
    for file in os.listdir(directory):
        if file.startswith("merged_") and file.endswith(".pdf"):
            try:
                os.remove(os.path.join(directory, file))

            except Exception as e:
                print(f"Warning: Could not remove old file: {e}")

    try:
        # PDF処理クラスのインスタンスを作成
        processor = PdfProcessor(directory)
        pdf_data = {}
        for file in sorted(os.listdir(directory)):
            if file.endswith(".pdf") and not file.startswith("merged_"):
                size, category, orientation = processor.get_pdf_info(file)
                pdf_data[file] = {'category': category, 'orientation': orientation}

        # PDFの結合処理を実行
        processor.merge_pdfs(pdf_data, hv_mode, start_size, end_size)
        print("PDF files merged successfully!")
    except Exception as e:
        print(f"An unexpected error occurred during merge: {e}")

    # スレッドの終了を通知
    event.set()

class CustomSpinner(BoxLayout):
    # kv側で values: 'A4', 'A3', ... と書けるようにするための定義
    values = ListProperty([])
    text = StringProperty('A4') # 初期値

class MyBoxLayout(BoxLayout):
    # フォルダ選択
    def select_directory(self):
        # tkinterのルートウィンドウがデスクトップに出ないように隠す
        root = tk.Tk()
        root.withdraw()
        
        # Windows環境でダイアログが背面に隠れるのを防ぐ
        if platform.system() == 'Windows':
            root.attributes('-topmost', True)
            
        try:
            # OS標準のフォルダ選択ダイアログを呼び出す
            path = filedialog.askdirectory()
            if path:
                self.ids.input_folder_path.text = path
        finally:
            root.destroy()

    def merge_pressed(self):
        work_dir = self.ids.input_folder_path.text
        hv_mode = self.ids.check_separate_vh.active
        # スピナーから値を取得
        start_size = self.ids.location_spinner_1.text
        end_size = self.ids.location_spinner_2.text

        print(start_size)
        print(end_size)
        
        if work_dir:
            self.ids.button_merge.disabled = True
            self.ids.button_merge.text = 'Running...'
            event = threading.Event()
            threading.Thread(target=merge_pdfs_main,
                             args=(work_dir, hv_mode, start_size, end_size, event)).start()
            
            def check_thread(dt):
                if event.is_set():
                    self.ids.button_merge.disabled = False
                    self.ids.button_merge.text = 'Merge'
                    return False
                return True

            Clock.schedule_interval(check_thread, 0.1)

class MyApp(App):
    title = f'{TITLE} v{VERSION}'
    Builder.load_file('style.kv')
    Window.size = (700, 250)
    Window.minimum_width, Window.minimum_height = 700,250
    Config.set('input', 'mouse', 'mouse,disable_multitouch')
    Config.set('kivy', 'exit_on_escape', '0')
    
    def build(self):
        return MyBoxLayout()

if __name__ == "__main__":
    MyApp().run()
