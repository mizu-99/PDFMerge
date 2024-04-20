import os
from PyPDF2 import PdfMerger, PdfReader

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

def get_page_sizes(directory, file_name):
    # ページサイズを格納するリストを初期化
    sizes = []
    # ファイルパスを生成
    pdf_path = os.path.join(directory, file_name)
    # PDFファイルをバイナリモードで開く
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)  # PDFリーダーを作成
        # 各ページについて処理
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]   # ページを取得
            sizes.append((page.mediabox[2] - page.mediabox[0], page.mediabox[3] - page.mediabox[1])) # ページのサイズを取得してリストに追加
    # 全てのページのサイズを返す
    return sizes

def delete_merged_files(directory, files):
    # ファイルを削除
    if not files:
        print("No merged files found.")
    else:
        for file in files:
            os.remove(os.path.join(directory, file))
        print("Merged files deleted successfully!")

def merge_pdfs():
    # 現在のディレクトリを取得
    work_dir = os.getcwd()

    # ファイルリストのインスタンスを作成
    file_manager = FileList(work_dir)
    # "merged_"が含まれるファイル名を取得
    files_delete = file_manager.get_file_by_name("merged_")

    # 結合済みファイルを削除
    delete_merged_files(work_dir, files_delete)

    # ファイル名のリストを取得
    files = file_manager.get_file_by_extension(".pdf")
    
    merger_a0 = PdfMerger()
    merger_a1 = PdfMerger()
    merger_a2 = PdfMerger()
    merger_a3 = PdfMerger()
    merger_a4 = PdfMerger()
    
    # ファイルをサイズごとに分類してマージ
    for file_name in files:
        sizes = get_page_sizes(work_dir, file_name)
        for size in sizes:
            if size[0] >= 1682 and size[1] >= 2378:  # A0
                merger_a0.append(file_name)
                break   # ページ数だけ重複結合されるのを防ぐためブレイク
            elif size[0] >= 1189 and size[1] >= 1682:  # A1
                merger_a1.append(file_name)
                break   # ページ数だけ重複結合されるのを防ぐためブレイク
            elif size[0] >= 841 and size[1] >= 1189:  # A2
                merger_a2.append(file_name)
                break   # ページ数だけ重複結合されるのを防ぐためブレイク
            elif size[0] >= 594 and size[1] >= 841:  # A3
                merger_a3.append(file_name)
                break   # ページ数だけ重複結合されるのを防ぐためブレイク
            else:  # A4以下
                merger_a4.append(file_name)
                break   # ページ数だけ重複結合されるのを防ぐためブレイク
        
    # 出力ファイル名を指定してマージ結果を保存
    if merger_a0.pages:
        merger_a0.write(os.path.join(work_dir, 'merged_A0.pdf'))
    if merger_a1.pages:
        merger_a1.write(os.path.join(work_dir, 'merged_A1.pdf'))
    if merger_a2.pages:
        merger_a2.write(os.path.join(work_dir, 'merged_A2.pdf'))
    if merger_a3.pages:
        merger_a3.write(os.path.join(work_dir, 'merged_A3.pdf'))
    if merger_a4.pages:
        merger_a4.write(os.path.join(work_dir, 'merged_A4.pdf'))
    
    print("PDF files merged successfully!")

if __name__ == "__main__":
    merge_pdfs()
