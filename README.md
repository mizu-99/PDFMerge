# PDF マージツール

![License-MIT](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)
![Python](https://custom-icon-badges.herokuapp.com/badge/Python-3572A5.svg?logo=Python&logoColor=white)
![SQL](https://custom-icon-badges.herokuapp.com/badge/SQL-e38c00.svg?logo=SQL&logoColor=white)

このスクリプトは、Python を使用して PDF ファイルをマージするためのユーティリティです。  
指定されたディレクトリ内の PDF ファイルを結合し、用紙サイズに応じて横向きと縦向きの結合を別々に行うことができます。

## 機能

- ディレクトリ内の PDF ファイルを以下の2種の方法で結合します。
  1. 用紙サイズ毎に結合を実施
  2. 用紙サイズ毎かつ横向きと縦向きの結合を別々に実施

## 依存関係

スクリプトを直接実施する場合は以下の依存ライブラリのインストールが必要です。

- Python 3
- PyPDF2

## 注意事項

- 対応サイズについて
  - A0 ~ A4 のA版サイズのみに対応
- スクリプト、exeファイルの保存場所について
  - 結合するPDFファイルと同一のディレクトリに保存する必要があります
- ファイル名称につて
  - 結合処理時に`merged`の名称が含まれるファイルが削除されます
  - `merged`が含まれるファイルを結合はできません

## 使用法

1. 結合した PDF ファイルと同じ場所にスクリプトを保存し、以下のコマンドで実行、もしくは`merge_pdfs.exe`ファイルを実行  

```shell
python merge_pdfs.py
```

2. 横向きと縦向きの PDF を別々に結合するかどうかを選択します。

```shell
横向きと縦向きのPDFを別々に結合しますか? (y/n): 
```

### y 選択の場合

- 横向きの PDF ファイルは、`merged_A0_H.pdf`, `merged_A1_H.pdf`, などという名前で結合・保存されます。
- 縦向きの PDF ファイルは、`merged_A0_V.pdf`, `merged_A1_V.pdf`, などという名前で結合・保存されます。

### n 選択の場合

- 横、縦向き関係なく PDF ファイルは、`merged_A0.pdf`, `merged_A1.pdf`, などという名前で結合・保存されます

## ライセンス

このプロジェクトは MIT ライセンスの下で利用できます。詳細については、[LICENSE](LICENSE) ファイルを参照してください。

[LICENSE]: https://github.com/mizu-99/PDFMerge/blob/master/LICENSE