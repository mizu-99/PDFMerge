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
    
    # ユーザーからの入力を受け取る
    while True:
        user_input = input("横向きと縦向きのPDFを別々に結合しますか? (y/n): ").strip().lower()
        if user_input == 'y' or user_input == 'n':
            break
        else:
            print("無効な入力です。'y'もしくは'n'を入力してください。")

    pdf_size = {
        'a0' : {'long' : 1189, 'short' : 841},
        'a1' : {'long' :  841, 'short' : 594},
        'a2' : {'long' :  594, 'short' : 420},
        'a3' : {'long' :  420, 'short' : 297},
        'a4' : {'long' :  297, 'short' : 210}
    }

    pdf_size_pt = {
        'a0' : {'long' : round(pdf_size['a0']['long'] / 25.4 * 72, 0), 'short' : round(pdf_size['a0']['short'] / 25.4 * 72, 1)},
        'a1' : {'long' : round(pdf_size['a1']['long'] / 25.4 * 72, 0), 'short' : round(pdf_size['a1']['short'] / 25.4 * 72, 1)},
        'a2' : {'long' : round(pdf_size['a2']['long'] / 25.4 * 72, 0), 'short' : round(pdf_size['a2']['short'] / 25.4 * 72, 1)},
        'a3' : {'long' : round(pdf_size['a3']['long'] / 25.4 * 72, 0), 'short' : round(pdf_size['a3']['short'] / 25.4 * 72, 1)},
        'a4' : {'long' : round(pdf_size['a4']['long'] / 25.4 * 72, 0), 'short' : round(pdf_size['a4']['short'] / 25.4 * 72, 1)}
    }

    threshold = 0.9

    if user_input == 'y':
        merger_a0_h = PdfMerger()
        merger_a0_v = PdfMerger()
        merger_a1_h = PdfMerger()
        merger_a1_v = PdfMerger()
        merger_a2_h = PdfMerger()
        merger_a2_v = PdfMerger()
        merger_a3_h = PdfMerger()
        merger_a3_v = PdfMerger()
        merger_a4_h = PdfMerger()
        merger_a4_v = PdfMerger()

        for file_name in files:
            sizes = get_page_sizes(work_dir, file_name)
            for size in sizes:
                # A0 横
                if (size[0] >= pdf_size_pt['a0']['long'] * threshold) and (size[1] >= pdf_size_pt['a0']['short'] * threshold):
                    merger_a0_h.append(file_name)
                # A1 縦
                elif (size[0] >= pdf_size_pt['a0']['short'] * threshold) and (size[1] >= pdf_size_pt['a0']['long'] * threshold):
                    merger_a0_v.append(file_name)
                # A1 横
                elif (size[0] >= pdf_size_pt['a1']['long'] * threshold) and (size[1] >= pdf_size_pt['a1']['short'] * threshold):
                    merger_a1_h.append(file_name)
                # A1 縦
                elif (size[0] >= pdf_size_pt['a1']['short'] * threshold) and (size[1] >= pdf_size_pt['a1']['long'] * threshold):
                    merger_a1_v.append(file_name)
                # A2 横
                elif (size[0] >= pdf_size_pt['a2']['long'] * threshold) and (size[1] >= pdf_size_pt['a2']['short'] * threshold):
                    merger_a2_h.append(file_name)
                # A2 縦
                elif (size[0] >= pdf_size_pt['a2']['short'] * threshold) and (size[1] >= pdf_size_pt['a2']['long'] * threshold):
                    merger_a2_v.append(file_name)
                # A3 横
                elif (size[0] >= pdf_size_pt['a3']['long'] * threshold) and (size[1] >= pdf_size_pt['a3']['short'] * threshold):
                    merger_a3_h.append(file_name)
                # A3 縦
                elif (size[0] >= pdf_size_pt['a3']['short'] * threshold) and (size[1] >= pdf_size_pt['a3']['long'] * threshold):
                    merger_a3_v.append(file_name)
                # A4 横
                elif (size[0] >= pdf_size_pt['a4']['long'] * threshold) and (size[1] >= pdf_size_pt['a4']['short'] * threshold):
                    merger_a4_h.append(file_name)
                # A4 縦
                elif (size[0] >= pdf_size_pt['a4']['short'] * threshold) and (size[1] >= pdf_size_pt['a4']['long'] * threshold):
                    merger_a4_v.append(file_name)
                else:
                    pass

                break   # ページ数だけ重複結合されるのを防ぐためブレイク

        # 出力ファイル名を指定してマージ結果を保存
        if merger_a0_h.pages:
            merger_a0_h.write(os.path.join(work_dir, 'merged_A0_H.pdf'))
        if merger_a0_v.pages:
            merger_a0_v.write(os.path.join(work_dir, 'merged_A0_V.pdf'))
        if merger_a1_h.pages:
            merger_a1_h.write(os.path.join(work_dir, 'merged_A1_H.pdf'))
        if merger_a1_v.pages:
            merger_a1_v.write(os.path.join(work_dir, 'merged_A1_V.pdf'))
        if merger_a2_h.pages:
            merger_a2_h.write(os.path.join(work_dir, 'merged_A2_H.pdf'))
        if merger_a2_v.pages:
            merger_a2_v.write(os.path.join(work_dir, 'merged_A2_V.pdf'))
        if merger_a3_h.pages:
            merger_a3_h.write(os.path.join(work_dir, 'merged_A3_H.pdf'))
        if merger_a3_v.pages:
            merger_a3_v.write(os.path.join(work_dir, 'merged_A3_V.pdf'))
        if merger_a4_h.pages:
            merger_a4_h.write(os.path.join(work_dir, 'merged_A4_H.pdf'))
        if merger_a4_v.pages:
            merger_a4_v.write(os.path.join(work_dir, 'merged_A4_V.pdf'))

    else:
        merger_a0 = PdfMerger()
        merger_a1 = PdfMerger()
        merger_a2 = PdfMerger()
        merger_a3 = PdfMerger()
        merger_a4 = PdfMerger()

        for file_name in files:
            sizes = get_page_sizes(work_dir, file_name)
            for size in sizes:
                # A0
                if (size[0] * size[1]) >= (pdf_size_pt['a0']['long'] * pdf_size_pt['a0']['short'] * threshold):
                    merger_a0.append(file_name)
                # A1
                elif (size[0] * size[1]) >= (pdf_size_pt['a1']['long'] * pdf_size_pt['a1']['short'] * threshold):
                    merger_a1.append(file_name)
                # A2
                elif (size[0] * size[1]) >= (pdf_size_pt['a2']['long'] * pdf_size_pt['a2']['short'] * threshold):
                    merger_a2.append(file_name)
                # A3
                elif (size[0] * size[1]) >= (pdf_size_pt['a3']['long'] * pdf_size_pt['a3']['short'] * threshold):
                    merger_a3.append(file_name)
                # A4
                elif (size[0] * size[1]) >= (pdf_size_pt['a4']['long'] * pdf_size_pt['a4']['short'] * threshold):
                    merger_a4.append(file_name)
                else:
                    pass
                
                break
        
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
