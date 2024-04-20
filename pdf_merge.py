import os
from PyPDF2 import PdfMerger, PdfReader

def get_page_sizes(pdf_path):
    sizes = []
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            sizes.append((page.mediabox[2] - page.mediabox[0],
                          page.mediabox[3] - page.mediabox[1]))
    return sizes

def merge_pdfs():
    # 現在のディレクトリを取得
    current_dir = os.getcwd()
    # ファイル名のリストを取得
    files = sorted([file for file in os.listdir(current_dir) if file.endswith('.pdf')])
    
    merger_a0 = PdfMerger()
    merger_a1 = PdfMerger()
    merger_a2 = PdfMerger()
    merger_a3 = PdfMerger()
    merger_a4 = PdfMerger()
    
    # ファイルをサイズごとに分類してマージ
    for file_name in files:
        sizes = get_page_sizes(file_name)
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
        merger_a0.write('merged_A0.pdf')
    if merger_a1.pages:
        merger_a1.write('merged_A1.pdf')
    if merger_a2.pages:
        merger_a2.write('merged_A2.pdf')
    if merger_a3.pages:
        merger_a3.write('merged_A3.pdf')
    if merger_a4.pages:
        merger_a4.write('merged_A4.pdf')
    
    print("PDF files merged successfully!")

if __name__ == "__main__":
    merge_pdfs()
