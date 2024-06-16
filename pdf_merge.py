import os
import threading
import platform
from pypdf import PdfMerger, PdfReader
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.lang import Builder
from kivy.core.window import Window
from kivy.config import Config
from kivy.core.text import LabelBase, DEFAULT_FONT  # フォント指定
from kivy.resources import resource_add_path        # フォント指定
from kivy.clock import Clock

if platform.system() == 'Linux':
    resource_add_path("/usr/share/fonts/opentype/noto/")
    LabelBase.register(DEFAULT_FONT, "NotoSansCJK-Regular.ttc")
# For Windows
else:
    resource_add_path("C:\Windows\Fonts")
    LabelBase.register(DEFAULT_FONT, "meiryo.ttc")

TITLE = 'PDFMerge'
VERSION = '0.2.5'

class FileList:
    def __init__(self, directory):
        # ディレクトリを設定
        self.directory = directory
    
    def get_file_by_name(self, keyword):
        # キーワードを含むファイルを取得
        files = [file for file in os.listdir(self.directory) if keyword in file]
        return sorted(files)
    
    def get_file_by_extension(self, extension):
        # 指定の拡張子のファイルを取得
        files = [file for file in os.listdir(self.directory) if file.endswith(extension)]
        return sorted(files)

def delete_merged_files(directory, files):
    # ファイルを削除
    if not files:
        print("No merged files found.")
    else:
        for file in files:
            os.remove(os.path.join(directory, file))
        print("Merged files deleted successfully!")

def convert_mm_to_pt(mm):
    return round(mm / 25.4 * 72.1, 1)

class PdfProcessor:
    def __init__(self, directory, threshold=0.9):
        self.directory  = directory
        self.threshold = threshold
        self.merges = {
            'A0'    : PdfMerger(),
            'A0_H'  : PdfMerger(),
            'A0_V'  : PdfMerger(),
            'A1'    : PdfMerger(),
            'A1_H'  : PdfMerger(),
            'A1_V'  : PdfMerger(),
            'A2'    : PdfMerger(),
            'A2_H'  : PdfMerger(),
            'A2_V'  : PdfMerger(),
            'A3'    : PdfMerger(),
            'A3_H'  : PdfMerger(),
            'A3_V'  : PdfMerger(),
            'A4'    : PdfMerger(),
            'A4_H'  : PdfMerger(),
            'A4_V'  : PdfMerger(),
            'error' : PdfMerger()
        }
        # self.merger_A0 = PdfMerger()
        
        self.pdf_sizes = {
        'A0' : {'long' : 1189, 'short' : 841},
        'A1' : {'long' :  841, 'short' : 594},
        'A2' : {'long' :  594, 'short' : 420},
        'A3' : {'long' :  420, 'short' : 297},
        'A4' : {'long' :  297, 'short' : 210}
        }
        
        self.pdf_size_pt = {size: {'long': convert_mm_to_pt(data['long']), 'short': convert_mm_to_pt(data['short'])} for size, data in self.pdf_sizes.items()}
    
    def get_pdf_info(self, file_name):
        # ページサイズを格納するリストを初期化
        sizes = []
        # ファイルパスを生成
        pdf_path = os.path.join(self.directory, file_name)
        
        # PDFファイルをバイナリモードで開く
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)  # PDFリーダーを作成
            # 各ページについて処理
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]   # ページを取得
                sizes.append((page.mediabox[2] - page.mediabox[0], page.mediabox[3] - page.mediabox[1])) # ページのサイズを取得してリストに追加
        
        # 複数ページのPDFの場合、最初のページのサイズだけを取得する
        size = list(sizes[:1][0])
        
        # PDFファイルの呼び番と縦横使いの情報’を判断
        # A0 横
        if (size[0] >= self.pdf_size_pt['A0']['long'] * self.threshold) and (size[1] >= self.pdf_size_pt['A0']['short'] * self.threshold):
            category    = 'A0'
            orientation = 'h'
        # A0 縦
        elif (size[0] >= self.pdf_size_pt['A0']['short'] * self.threshold) and (size[1] >= self.pdf_size_pt['A0']['long'] * self.threshold):
            category    = 'A0'
            orientation = 'v'
        # A1 横
        elif (size[0] >= self.pdf_size_pt['A1']['long'] * self.threshold) and (size[1] >= self.pdf_size_pt['A1']['short'] * self.threshold):
            category    = 'A1'
            orientation = 'h'
        # A1 縦
        elif (size[0] >= self.pdf_size_pt['A1']['short'] * self.threshold) and (size[1] >= self.pdf_size_pt['A1']['long'] * self.threshold):
            category    = 'A1'
            orientation = 'v'
        # A2 横
        elif (size[0] >= self.pdf_size_pt['A2']['long'] * self.threshold) and (size[1] >= self.pdf_size_pt['A2']['short'] * self.threshold):
            category    = 'A2'
            orientation = 'h'
        # A2 縦
        elif (size[0] >= self.pdf_size_pt['A2']['short'] * self.threshold) and (size[1] >= self.pdf_size_pt['A2']['long'] * self.threshold):
            category    = 'A2'
            orientation = 'v'
        # A3 横
        elif (size[0] >= self.pdf_size_pt['A3']['long'] * self.threshold) and (size[1] >= self.pdf_size_pt['A3']['short'] * self.threshold):
            category    = 'A3'
            orientation = 'h'
        # A3 縦
        elif (size[0] >= self.pdf_size_pt['A3']['short'] * self.threshold) and (size[1] >= self.pdf_size_pt['A3']['long'] * self.threshold):
            category    = 'A3'
            orientation = 'v'
        # A4 横
        elif (size[0] >= self.pdf_size_pt['A4']['long'] * self.threshold) and (size[1] >= self.pdf_size_pt['A4']['short'] * self.threshold):
            category    = 'A4'
            orientation = 'h'
        # A4 縦
        elif (size[0] >= self.pdf_size_pt['A4']['short'] * self.threshold) and (size[1] >= self.pdf_size_pt['A4']['long'] * self.threshold):
            category    = 'A4'
            orientation = 'v'
        # 例外
        else:
            category    = 'Unknown'
            orientation = 'Unknown'
        
        # 最初のページのサイズを返す
        return size, category, orientation
    
    def merge_pdf(self, pdf_data):
        for file_name in pdf_data.keys():
            file_path = os.path.join(self.directory, file_name)
            # A0
            if pdf_data[file_name]['category'] == 'A0':
                self.merges['A0'].append(file_path)
            # A1
            elif pdf_data[file_name]['category'] == 'A1':
                self.merges['A1'].append(file_path)
            # A2
            elif pdf_data[file_name]['category'] == 'A2':
                self.merges['A2'].append(file_path)
            # A3
            elif pdf_data[file_name]['category'] == 'A3':
                self.merges['A3'].append(file_path)
            # A4
            elif pdf_data[file_name]['category'] == 'A4':
                self.merges['A4'].append(file_path)
            # Error
            else:
                self.merges['error'].append(file_path)
        
        # 出力ファイル名を指定してマージ結果を保存
        # A0
        if self.merges['A0'].pages:
            self.merges['A0'].write(os.path.join(self.directory, 'merged_A0.pdf'))
        # A1
        if self.merges['A1'].pages:
            self.merges['A1'].write(os.path.join(self.directory, 'merged_A1.pdf'))
        # A2
        if self.merges['A2'].pages:
            self.merges['A2'].write(os.path.join(self.directory, 'merged_A2.pdf'))
        # A3
        if self.merges['A3'].pages:
            self.merges['A3'].write(os.path.join(self.directory, 'merged_A3.pdf'))
        # A4
        if self.merges['A4'].pages:
            self.merges['A4'].write(os.path.join(self.directory, 'merged_A4.pdf'))
        # Error
        if self.merges['error'].pages:
            self.merges['error'].write(os.path.join(self.directory, 'merged_error.pdf'))
    
    def merge_pdf_hv(self, pdf_data):
        for file_name in pdf_data.keys():
            file_path = os.path.join(self.directory, file_name)
            # A0_H
            if pdf_data[file_name]['category'] == 'A0' and pdf_data[file_name]['orientation'] == 'h':
                self.merges['A0_H'].append(file_path)
            # A0_V
            elif pdf_data[file_name]['category'] == 'A0' and pdf_data[file_name]['orientation'] == 'v':
                self.merges['A0_V'].append(file_path)
            # A1_H
            elif pdf_data[file_name]['category'] == 'A1' and pdf_data[file_name]['orientation'] == 'h':
                self.merges['A1_H'].append(file_path)
            # A1_V
            elif pdf_data[file_name]['category'] == 'A1' and pdf_data[file_name]['orientation'] == 'v':
                self.merges['A1_V'].append(file_path)
            # A2_H
            elif pdf_data[file_name]['category'] == 'A2' and pdf_data[file_name]['orientation'] == 'h':
                self.merges['A2_H'].append(file_path)
            # A2_V
            elif pdf_data[file_name]['category'] == 'A2' and pdf_data[file_name]['orientation'] == 'v':
                self.merges['A2_V'].append(file_path)
            # A3_H
            elif pdf_data[file_name]['category'] == 'A3' and pdf_data[file_name]['orientation'] == 'h':
                self.merges['A3_H'].append(file_path)
            # A3_V
            elif pdf_data[file_name]['category'] == 'A3' and pdf_data[file_name]['orientation'] == 'v':
                self.merges['A3_V'].append(file_path)
            # A4_H
            elif pdf_data[file_name]['category'] == 'A4' and pdf_data[file_name]['orientation'] == 'h':
                self.merges['A4_H'].append(file_path)
            # A4_V
            elif pdf_data[file_name]['category'] == 'A4' and pdf_data[file_name]['orientation'] == 'v':
                self.merges['A4_V'].append(file_path)
            # Error
            else:
                self.merges['error'].append(file_path)
        
        # 出力ファイル名を指定してマージ結果を保存
        # A0_H
        if self.merges['A0_H'].pages:
            self.merges['A0_H'].write(os.path.join(self.directory, 'merged_A0_H.pdf'))
        # A0_V
        if self.merges['A0_V'].pages:
            self.merges['A0_V'].write(os.path.join(self.directory, 'merged_A0_V.pdf'))
        # A1_H
        if self.merges['A1_H'].pages:
            self.merges['A1_H'].write(os.path.join(self.directory, 'merged_A1_H.pdf'))
        # A1_V
        if self.merges['A1_V'].pages:
            self.merges['A1_V'].write(os.path.join(self.directory, 'merged_A1_V.pdf'))
        # A2_H
        if self.merges['A2_H'].pages:
            self.merges['A2_H'].write(os.path.join(self.directory, 'merged_A2_H.pdf'))
        # A2_V
        if self.merges['A2_V'].pages:
            self.merges['A2_V'].write(os.path.join(self.directory, 'merged_A2_V.pdf'))
        # A3_H
        if self.merges['A3_H'].pages:
            self.merges['A3_H'].write(os.path.join(self.directory, 'merged_A3_H.pdf'))
        # A3_V
        if self.merges['A3_V'].pages:
            self.merges['A3_V'].write(os.path.join(self.directory, 'merged_A3_V.pdf'))
        # A4_H
        if self.merges['A4_H'].pages:
            self.merges['A4_H'].write(os.path.join(self.directory, 'merged_A4_H.pdf'))
        # A4_V
        if self.merges['A4_V'].pages:
            self.merges['A4_V'].write(os.path.join(self.directory, 'merged_A4_V.pdf'))
        # Error
        if self.merges['error'].pages:
            self.merges['error'].write(os.path.join(self.directory, 'merged_error.pdf'))

def merge_pdfs(directory, merge_status, event):
    # 現在のディレクトリを取得
    work_dir = directory
    
    # ファイルリストのインスタンスを作成
    file_manager = FileList(work_dir)
    # "merged_"が含まれるファイル名を取得
    files_delete = file_manager.get_file_by_name("merged_")
    
    # 結合済みファイルを削除
    delete_merged_files(work_dir, files_delete)
    
    # 走査ディレクトリの中かからpdfファルを検索
    files = file_manager.get_file_by_extension(".pdf")
    
    # PdfPProcessorのインスタンス化
    processor    = PdfProcessor(directory)
    
    # 取得したPDFファイルのデータベース作成
    pdf_data = {}
    for file in files:
        size, category, orientation = processor.get_pdf_info(file)  # PDFファイルの情報を
        # PDFファイルの情報を記録
        pdf_data[file] = {
            'size'        : '{} x {}'.format(round(size[0], 1), round(size[1], 1)),
            'category'    : category,
            'orientation' : orientation
        }
    
    if merge_status == True:
        processor.merge_pdf_hv(pdf_data)
    else:
        processor.merge_pdf(pdf_data)
    
    print("PDF files merged successfully!")

    event.set() # スレッドの終了を通知

class MyBoxLayout(BoxLayout):
    def submit_pressed(self):
        input_text = self.ids.input_folder_path.text
        print("FOLDER PATH:", input_text)
    
    def merge_pressed(self):
        work_dir             = self.ids.input_folder_path.text
        check_separate_vh    = self.ids.check_separate_vh.active
        check_no_separate_vh = self.ids.check_no_separate_vh.active
    
        if check_separate_vh == True and check_no_separate_vh == False:
            merge_status = True
        elif check_separate_vh == False and check_no_separate_vh == True:
            merge_status = False
        else:
            self.ids.check_separate_vh.active = True # Treuのチェックボックスが0のときは初期状態に設定
            merge_status = True
    
        if work_dir:
            self.ids.button_merge.disabled = True # 「Merge」ボタンを無効化
            self.ids.button_merge.text = 'Running...'
            print('Running')
            event = threading.Event()
            merge_thread = threading.Thread(target= merge_pdfs, args=(work_dir, merge_status, event, ))
            merge_thread.start()
            # merge_pdfs(work_dir, merge_status)
            # スレッドが終了するのを待つ
            def check_thread(dt):
                if event.is_set():
                    self.ids.button_merge.disabled = False # 「Merge」ボタンを無効化
                    self.ids.button_merge.text = 'Merge'
                    return False # スケジューリングを停止
                return True # スケジューリングを継続
        
            Clock.schedule_interval(check_thread, 0.1) # 0.1秒ごとにスレッドの状態をチェック
        else:
            pass

class MyApp(App):
    title = '{} v{}'.format(TITLE, VERSION)
    # kvファイルの読み込み
    Builder.load_file('style.kv')
    
    # windowサイズの指定
    window_width  = 700
    window_height = 100
    Window.size   = (window_width, window_height)
    
    # 指定のサイズ以下へのウィンドウのサイズ変更を不可にする
    Window.minimum_width  = window_width
    Window.minimum_height = window_height
    
    # guiの振舞いの設定
    Config.set('input', 'mouse', 'mouse,disable_multitouch')  # 右クリック赤丸消去
    Config.set('kivy', 'exit_on_escape', '0')  # kivy ESC無効
    
    def build(self):
        return MyBoxLayout()

if __name__ == "__main__":
    MyApp().run()
