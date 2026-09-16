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
    argparser.add_argument('-u',
                           '--update_pages',
                           action="store_true",
                           help='Update only page numbers in existing filelist.csv')
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

def merge_with_existing(new_df, existing_csv_path):
    """
    既存のCSVと新しいDataFrameをマージし、既存のデータと順序を保持する。

    :param new_df: 新しいファイル情報を格納した DataFrame
    :param existing_csv_path: 既存のCSVファイルパス
    :return: マージされた DataFrame
    """
    if not os.path.exists(existing_csv_path):
        return new_df

    existing_df = pd.read_csv(existing_csv_path)

    # 既存に있는 파일은 기존 순서 유지
    existing_files = existing_df['filename'].tolist()
    new_files = new_df['filename'].tolist()

    # 既存にあるファイルと新しいファイルに分ける
    existing_in_new = [f for f in existing_files if f in new_files]
    new_added = [f for f in new_files if f not in existing_files]

    # 既存順序を基準に、既存ファイルと新しいファイルを結合
    ordered_files = existing_in_new + new_added
    new_df['order'] = new_df['filename'].apply(lambda x: ordered_files.index(x))
    new_df = new_df.sort_values('order').drop(columns=['order']).reset_index(drop=True)

    # filenameをキーにしてマージ
    merged_df = new_df.merge(
        existing_df[['filename', 'title', 'author', 'toc_text', 'add_pagenum', 'start_page']],
        on='filename',
        how='left',
        suffixes=('_new', '_old')
    )

    # 既存データがあれば使用、なければ新しいデータを使用
    for col in ['title', 'author', 'toc_text', 'add_pagenum']:
        if f'{col}_old' in merged_df.columns:
            merged_df[col] = merged_df[f'{col}_old'].fillna(merged_df[f'{col}_new'])
            merged_df.drop(columns=[f'{col}_old', f'{col}_new'], inplace=True)

    # start_page カラムがなければ空値で埋める
    if 'start_page' not in merged_df.columns:
        merged_df['start_page'] = pd.NA

    return merged_df

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

    # --update_pages フラグが有効な場合、既存のCSVと統合
    if args.update_pages:
        file_table = merge_with_existing(file_table, filelist)
        print(f"既存のCSVとマージしました。ページ番号を更新します。")

    file_table.to_csv(filelist, index=False, encoding="utf-8-sig")
    print(f"CSVファイルを作成しました: {filelist} (ファイル数: {len(file_table)})")



