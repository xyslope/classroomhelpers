from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_TAB_ALIGNMENT, WD_TAB_LEADER
import sys
import os
import io
import shutil
from pathlib import Path
from argparse import ArgumentParser
import pandas as pd

import comtypes.client

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
    argparser.add_argument('-t',
                           '--tocfile',
                           type=str,
                           default='index.docx',
                           help='Table Of Contents file')
    argparser.add_argument('-pc',
                           '--page_count_from',
                           type=int,
                           default=0,
                           help='Count Page from x. First document is 0.')
    return argparser.parse_args()

def add_toc_entry(doc, toc_text, page_number):
    """
    Word の目次エントリーを追加する。
    ページ番号が None の場合は、リーダー（点線）とページ番号を表示しない。

    :param doc: Word 文書オブジェクト
    :param toc_text: 目次のテキスト
    :param page_number: ページ番号（None の場合は表示しない）
    """
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT  # 左揃え

    # タブストップの設定（右端にページ番号を配置し、リーダーを挿入）
    tab_stops = paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Inches(6), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)  # 右揃え + 点々リーダー

    # 目次テキストを追加
    run = paragraph.add_run(toc_text)
    run.font.size = Pt(12)

    if page_number is not None:
        # タブを挿入（リーダーを追加）
        paragraph.add_run("\t")  # タブ文字でリーダーを挿入

        # ページ番号を追加（右端）
        run_page = paragraph.add_run(str(page_number))
        run_page.font.size = Pt(12)

def compute_toc_pages(tocfile):
    """
    Word の `ComputeStatistics(2)` を使用して、目次のページ数を取得する。

    :param tocfile: Word ファイルのパス
    :return: 目次のページ数
    """
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = False  # 非表示で実行

    doc = word.Documents.Open(tocfile)
    total_pages = doc.ComputeStatistics(2)  # wdStatisticPages = 2 でページ数を取得
    doc.Close()
    word.Quit()

    return total_pages


def create_1index(filelist, tocfile, page_count_from=0):
    """
    CSV からファイルリストを読み込み、目次を作成する。
    :param filelist_csv: 目次を作成するための CSV ファイル
    :param tocfile: 出力する目次の Word ファイル
    :param page_count_from: どこから「1」とカウントするか（デフォルトは 1）
    """

    doc = Document()
    doc.add_heading("目次", level=1)

    current_page = 1  # まずは 1 からカウント
    for i, entry in filelist.iterrows():
        if i >= page_count_from:  # ✅ 指定行（`page_count_from`）からカウント開始
            add_toc_entry(doc, entry["toc_text"], current_page)
            current_page += entry["page_count"]
        else:
            add_toc_entry(doc, entry["toc_text"], None)  # 指定行より前はページ番号なし
    doc.save(tocfile)

def create_index(filelist, tocfile, page_count_from=0):
    """
    CSV からファイルリストを読み込み、目次作成関数を呼ぶ
    目次のページ数を取得し、正しいページ番号で目次を再作成する。
    
    :param filelist_csv: 目次を作成するための CSV ファイル
    :param tocfile: 出力する目次の Word ファイル
    :param page_count_from: どこから「1」とカウントするか（デフォルトは 1）
    """

    # 1️⃣ 仮の目次を作成（ページ番号の取得用）
    doc1 = create_1index(filelist,tocfile, page_count_from)

    # 2️⃣ 目次のページ数を取得し、filelist を更新
    toc_pages = compute_toc_pages(tocfile)
    print(f"目次のページ数: {toc_pages}")

    # `filelist` のページ数を再計算
    current_page = page_count_from + toc_pages  # 目次の次のページからカウント開始
    updated_filelist = filelist.copy()
    for index, entry in updated_filelist.iterrows():
        updated_filelist.at[index, "start_page"] = current_page  # 正しい開始ページを記録
        current_page += entry["page_count"]
# updated_filelist.to_csv(filelist, index=False, encoding="utf-8-sig")

    # 3️⃣ 正しいページ番号で目次を再作成
    doc2 = create_1index(updated_filelist,tocfile, page_count_from)
    print(f"目次を最終更新しました: {tocfile} (目次ページ数: {toc_pages})")
    return updated_filelist

if __name__ == "__main__":
    # get_options
    args = get_options()
    source_path = args.source_path
    docdir = Path(source_path) / args.doc_dir.lstrip("\\/")
    if not docdir.exists():
        print('ソースディレクトリが存在しません。')
        sys.exit()
 
    # create filelist
    #    filelist = pd.read_csv(Path(source_path) / 'filelist.csv')
    filelist = pd.read_csv(str(Path(source_path) / args.filelist), encoding="utf-8-sig")
    tocfile = str(Path(docdir) / args.tocfile)
    newfilelist = create_index(filelist, tocfile, args.page_count_from)
    newfilelist.to_csv(str(Path(source_path) / args.filelist), encoding="utf-8-sig")
