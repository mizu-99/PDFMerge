import os
import threading
import platform
from pypdf import PdfWriter, PdfReader
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.lang import Builder
from kivy.core.window import Window
from kivy.config import Config
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.resources import resource_add_path
from kivy.clock import Clock

if platform.system() == 'Linux':
    resource_add_path("/usr/share/fonts/opentype/noto/")
    LabelBase.register(DEFAULT_FONT, "NotoSansCJK-Regular.ttc")
else:
    resource_add_path("C:\\Windows\\Fonts")
    LabelBase.register(DEFAULT_FONT, "meiryo.ttc")

TITLE = 'PDFMerge'
VERSION = '0.4.0'

PDF_SIZES_MM = {
    'A0': (1189, 841), 'A1': (841, 594), 'A2': (594, 420),
    'A3': (420, 297), 'A4': (297, 210)
}

def convert_mm_to_pt(mm):
    return round(mm / 25.4 * 72, 1)

PDF_SIZES_PT = {k: (convert_mm_to_pt(v[0]), convert_mm_to_pt(v[1])) for k, v in PDF_SIZES_MM.items()}

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
        self.merges = {key: PdfWriter() for key in [
            'A0', 'A0_H', 'A0_V', 'A1', 'A1_H', 'A1_V', 'A2', 'A2_H', 'A2_V',
            'A3', 'A3_H', 'A3_V', 'A4', 'A4_H', 'A4_V', 'AboveA3', 'error'
        ]}

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
    # all_merge_above_a3: A3以上のサイズを全結合(縦横の判定をしない)
    def merge_pdfs(self, pdf_data, hv_mode=True, all_merge_above_a3=False):
        sorted_list = sorted(pdf_data.items(), key=lambda x: x[0])

        for file_name, meta in sorted_list:
                    category = meta['category']
                    orientation = meta['orientation'].upper()

                    # A3以上全結合モードが有効で、対象がA3以上（A0, A1, A2, A3）の場合
                    if all_merge_above_a3 and category in ['A0', 'A1', 'A2', 'A3']:
                        key = 'AboveA3'
                    # 従来の縦横分けモード
                    elif hv_mode:
                        key = f"{category}_{orientation}"
                    # 縦横を分けないモード
                    else:
                        key = category

                    self.merges.get(key, self.merges['error']).append(os.path.join(self.directory, file_name))

        for key, merger in self.merges.items():
            if merger.pages:
                output_path = os.path.join(self.directory, f'merged_{key}.pdf')
                if os.path.exists(output_path):
                    os.remove(output_path)
                merger.write(output_path)


# PDF結合処理のメイン関数
def merge_pdfs(directory, hv_mode, all_merge_above_a3, event):
    # 既存の結合ファイルを削除
    # 既存の結合済みPDFを削除
    keys = ['A0', 'A0_H', 'A0_V', 'A1', 'A1_H', 'A1_V', 'A2', 'A2_H', 'A2_V', 'A3', 'A3_H', 'A3_V', 'A4', 'A4_H', 'A4_V', 'AboveA3', 'error']
    for key in keys:
        output_path = os.path.join(directory, f'merged_{key}.pdf')
        if os.path.exists(output_path):
            os.remove(output_path)
    # PDF処理クラスのインスタンスを作成
    processor = PdfProcessor(directory)
    pdf_data = {}
    for file in sorted(os.listdir(directory)):
        if file.endswith(".pdf"):
            size, category, orientation = processor.get_pdf_info(file)
            pdf_data[file] = {'size': f'{size[0]:.1f} x {size[1]:.1f}', 'category': category, 'orientation': orientation}
    # PDFの結合処理を実行
    processor.merge_pdfs(pdf_data, hv_mode, all_merge_above_a3)
    print("PDF files merged successfully!")
    # スレッドの終了を通知
    event.set()

class MyBoxLayout(BoxLayout):
    def merge_pressed(self):
        work_dir           = self.ids.input_folder_path.text
        hv_mode            = self.ids.check_separate_vh.active
        all_merge_above_a3 = self.ids.check_all_merge_above_A3.active
        
        if work_dir:
            self.ids.button_merge.disabled = True
            self.ids.button_merge.text = 'Running...'
            event = threading.Event()
            threading.Thread(target=merge_pdfs, args=(work_dir, hv_mode, all_merge_above_a3, event)).start()
            
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
    Window.size = (500, 130)
    Window.minimum_width, Window.minimum_height = 500, 130
    Config.set('input', 'mouse', 'mouse,disable_multitouch')
    Config.set('kivy', 'exit_on_escape', '0')
    
    def build(self):
        return MyBoxLayout()

if __name__ == "__main__":
    MyApp().run()
