import os
import threading
import platform
from pypdf import PdfMerger, PdfReader
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
VERSION = '0.2.7'

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
        self.merges = {key: PdfMerger() for key in [
            'A0', 'A0_H', 'A0_V', 'A1', 'A1_H', 'A1_V', 'A2', 'A2_H', 'A2_V',
            'A3', 'A3_H', 'A3_V', 'A4', 'A4_H', 'A4_V', 'error'
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
    # pdf_data: 各PDFのサイズ・カテゴリ・向きの情報
    # hv_mode: Trueの場合は横向き・縦向きを分けて処理
    def merge_pdfs(self, pdf_data, hv_mode=True):
        sorted_list = sorted(pdf_data.items(), key=lambda x: x[0])
        if hv_mode:
            sorted_list = sorted(sorted_list, reverse=True, key=lambda x: x[1]['orientation'].upper() == 'H')
        for file_name, meta in sorted_list:
            key = meta['category'] if not hv_mode else f"{meta['category']}_{meta['orientation'].upper()}"
            self.merges.get(key, self.merges['error']).append(os.path.join(self.directory, file_name))

        for key, merger in self.merges.items():
            output_path = os.path.join(self.directory, f'merged_{key}.pdf')
            if os.path.exists(output_path):
                os.remove(output_path)
            if merger.pages:
                merger.write(os.path.join(self.directory, f'merged_{key}.pdf'))


# PDF結合処理のメイン関数
def merge_pdfs(directory, hv_mode, event):
    # 既存の結合ファイルを削除
    # 既存の結合済みPDFを削除
    for key in ['A0', 'A0_H', 'A0_V', 'A1', 'A1_H', 'A1_V', 'A2', 'A2_H', 'A2_V', 'A3', 'A3_H', 'A3_V', 'A4', 'A4_H', 'A4_V', 'error']:
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
    processor.merge_pdfs(pdf_data, hv_mode)
    print("PDF files merged successfully!")
    # スレッドの終了を通知
    event.set()

class MyBoxLayout(BoxLayout):
    def merge_pressed(self):
        work_dir = self.ids.input_folder_path.text
        hv_mode = self.ids.check_separate_vh.active
        
        if work_dir:
            self.ids.button_merge.disabled = True
            self.ids.button_merge.text = 'Running...'
            event = threading.Event()
            threading.Thread(target=merge_pdfs, args=(work_dir, hv_mode, event)).start()
            
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
    Window.size = (700, 100)
    Window.minimum_width, Window.minimum_height = 700, 100
    Config.set('input', 'mouse', 'mouse,disable_multitouch')
    Config.set('kivy', 'exit_on_escape', '0')
    
    def build(self):
        return MyBoxLayout()

if __name__ == "__main__":
    MyApp().run()
