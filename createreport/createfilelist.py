import sys
import os
import io
import shutil
import comtypes.client
from pathlib import Path
from argparse import ArgumentParser
import pandas as pd
import csv

def get_options():
    argparser = ArgumentParser()
    argparser.add_argument('-s',
                           '--source_path',
                           type=str,
                           default=os.getcwd(),
                           help='Document root')
    argparser.add_argument('-dd',
                           '--doc_dir',
                           type=str,
                           default='docs',
                           help='Path to doc directory')
    argparser.add_argument('-l',
                           '--filelist',
                           type=str,
                           default='filelist.csv',
                           help='File list')
    argparser.add_argument('-b',
                           '--if_addblank',
                           action="store_true",
                           help='If number of pages is odd, add a blank page.')
    return argparser.parse_args()

def create_filelist(docs_folder):
    # フォルダ内のファイル一覧を取得
    file_list = [
    f for f in os.listdir(docs_folder)
    if f.endswith(".docx") and not f.startswith("~$")
    ]

    # データフレームを作成
    df = pd.DataFrame(file_list, columns=["filename"])
      
    # 各カラムを設定
    # filename から拡張子を除き、title に設定
    df["title"] = df["filename"].apply(lambda x: os.path.splitext(x)[0])
    df["author"] = "ごんべえ"  # author は固定値
    df["toc_text"] = df.apply(lambda row: f"{row['title']} ({row['author']})", axis=1)
    df["add_pagenum"] = True  # add_pagenum = True

    return df  # DataFrame を返す

def add_pages(file_table, docs_folder, if_add_blank):
    """
    各ファイルのページ数を計算し、file_table に 'page_count' カラムを追加する。

    :param file_table: ファイル情報を格納した Pandas DataFrame
    :param docs_folder: ファイルが保存されているフォルダ
    :return: ページ数を追加した DataFrame
    """
    # Word アプリケーションを起動
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # 非表示で実行

    # 'page_count' カラムを追加（初期値は 0）
    file_table["page_count"] = 0

    # 各ファイルのページ数を取得
    for index, row in file_table.iterrows():
        file_path = os.path.join(docs_folder, row["filename"])  # フルパスを作成
        try:
            doc = word.Documents.Open(file_path)  # Word 文書を開く
            total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2
            file_table.at[index, "page_count"] = total_pages  # DataFrame に記録
            if if_add_blank:
                file_table["page_count"] = file_table["page_count"].apply(lambda x: x + 1 if x % 2 != 0 else x)
            doc.Close()
        except Exception as e:
            print(f"ページ数の取得に失敗しました: {row['filename']} - {e}")
    
    # Word を終了
    word.Quit()
    
    return file_table

if __name__ == "__main__":
    # get_options
    args = get_options()
    source_path = args.source_path
    docdir = Path(source_path) / args.doc_dir.lstrip("\\/")
    if not docdir.exists():
        print('ソースディレクトリが存在しません。')
        sys.exit()
    filelist =  str(Path(source_path) / args.filelist)
 
    file_table = create_filelist(docdir)
    file_table = add_pages(file_table, docdir, args.if_addblank)

    file_table.to_csv(filelist, index=False, encoding="utf-8-sig")
    print(f"CSVファイルを作成しました: {filelist} (ファイル数: {len(file_table)})")



