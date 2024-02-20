import sys
import os
import shutil
import comtypes.client
import glob
import pathlib
import PyPDF2
from PyPDF2 import PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4, portrait
from argparse import ArgumentParser
import pandas as pd
import io


# ページ番号の下からの位置
PAGE_BOTTOM = 10 * mm
# ページ番号のプレフィックス
PAGE_PREFIX = "資料 - "
# フォント登録
pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))

# オプション引数を受け取る
#（出所：「Pythonでオプション引数を受け取る - Qiita」、https://qiita.com/taashi/items/400871fb13df476f42d2  ）

def get_options():
    argparser = ArgumentParser()
    argparser.add_argument('-b', '--add_blank',
                           action="store_true",
                           help='If number of pages is odd, add a blank page.')
    argparser.add_argument('-sc', '--skip_convert', 
                           help='Skip convert from docx to pdf.',
                           action="store_true",
                           )
    argparser.add_argument('-w', '--wipe_tempdir',
                           action="store_true",
                           help='Wipe working directory.')
    argparser.add_argument('-s', '--source_path', type=str,
                           default = os.getcwd(),
                           help='Document root')
    argparser.add_argument('-dd', '--doc_dir', type=str,
                           default = 'docs',
                           help='Path to docs')
    argparser.add_argument('-o', '--out_file', type=str,
                           default = 'out.pdf',
                           help='Path to docs')
    argparser.add_argument('-td', '--temp_dir', type=str,
                           default = 'temp',
                           help='Working directory')
    argparser.add_argument('-ww', '--wipe_workingfiles',
                           action = "store_true",
                           help='wipe working files')
    argparser.add_argument('-p', '--from_pagenum', type=int,
                           default = 0,
                           help='Add page num from page x')
    return argparser.parse_args()

def convert(in_file, out_file):
    """convert word file to pdf
    in_file: word file with fullpath
    out_file: pdf file with fullpath
    """
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=17)
    doc.Close()
    word.Quit()

def pdf_merger(out_pdf, pdfs, add_blank):
    """create a combined pdf file.
    out_pdf: combined documents
    pdfs: pdf files with full path.
    tmpfile: temp file for working.
    add_blank: if true, add blank page to a pdf files with odd page numbers.
    """
    pdfindex = {}
    print('Merging your documents...')
    merger = PyPDF2.PdfMerger()
    newpage = 1
    for pdf in pdfs:
        pdfindex[os.path.basename(pdf)] = newpage
        merger.append(pdf)
        if add_blank & (len(merger.pages) % 2) == 1:
            merger.append(blankpage)
        newpage = len(merger.pages)
        print(pdf, '(pages: ', len(merger.pages), ')')
    for key, val in pdfindex.items():
        if not key == '-':
            merger.add_outline_item(key, val, parent=None)
    merger.write(out_pdf)
    merger.close()
    return pdfindex


def add_page_number(input_file: str, output_file: str, start_num: int = 1, record_from: int = 0):
    """
    既存PDFにページ番号を追加する
    """
    # 既存PDF（ページを付けるPDF）
    fi = open(input_file, 'rb')
    pdf_reader = PyPDF2.PdfReader(fi)
    pages_num = len(pdf_reader.pages)

    # ページ番号を付けたPDFの書き込み用
    pdf_writer = PyPDF2.PdfWriter()

    # ページ番号だけのPDFをメモリ（binary stream）に作成
    bs = io.BytesIO()
    c = canvas.Canvas(bs)
    for i in range(0, pages_num):
        # 既存PDF
        pdf_page = pdf_reader.pages[i]
        # PDFページのサイズ
        page_size = get_page_size(pdf_page)
        # ページ番号のPDF作成
        create_page_number_pdf(c, page_size, i + start_num)
    c.save()

    # ページ番号だけのPDFをメモリから読み込み（seek操作はPyPDF2に実装されているので不要）
    pdf_num_reader = PyPDF2.PdfReader(bs)

    # 既存PDFに１ページずつページ番号を付ける
    for i in range(0, pages_num):
        # 既存PDF
        pdf_page = pdf_reader.pages[i]
        # ページ番号だけのPDF
        pdf_num = pdf_num_reader.pages[i]
        print("ページを追加：" + str(record_from))
        if i >= record_from:
            # ２つのPDFを重ねる
            pdf_page.merge_page(pdf_num)
        pdf_writer.add_page(pdf_page)

    # ページ番号を付けたPDFを保存
    fo = open(output_file, 'wb')
    pdf_writer.write(fo)

    bs.close()
    fi.close()
    fo.close()


def create_page_number_pdf(c: canvas.Canvas, page_size: tuple, page_num: int):
    """
    ページ番号だけのPDFを作成
    """
    c.setPageSize(page_size)
    c.setFont("HeiseiKakuGo-W5", 10)
    if page_num % 2 != 0:
        c.drawString(25,
                      PAGE_BOTTOM,
                      "-" + str(page_num) + "-")
    else:
        c.drawRightString(page_size[0] - 25,
                      PAGE_BOTTOM,
                      "-" + str(page_num) + "-")
    c.showPage()


def get_page_size(page: PageObject) -> tuple:
    """
    既存PDFからページサイズ（幅, 高さ）を取得する
    """
    page_box = page.mediabox
    width = page_box.right - page_box.left
    height = page_box.top - page_box.bottom

    return float(width), float(height)

def add_outline(infile, outfile, pdfindex):
    reader = PyPDF2.PdfReader(infile)
    writer = PyPDF2.PdfWriter()
    writer.append_pages_from_reader(reader)
    #    writer.clone_document_from_reader(reader)
    for key, val in pdfindex.items():
        if not key == '-':
            print(key, val)
            writer.add_outline_item(key, val, parent=None)  # add bookmark
    with open(outfile, "wb") as fp:
        writer.write(fp)

if __name__ == "__main__":
    args = get_options()
    source_path = args.source_path + '\\'
    docdir = source_path + args.doc_dir + '\\'

    # ファイル読み込み
    if not os.path.exists(docdir):
        print('ソースディレクトリが存在しません。')
        sys.exit()

    filelist = pd.read_csv(source_path + '\\filelist.csv')
    filelist = filelist.set_index('file')
    tmpdir = source_path  + args.temp_dir + '\\'
                    
    if not args.skip_convert:
        if os.path.exists(tmpdir): shutil.rmtree(tmpdir)
        os.mkdir(tmpdir)

    # A4の新規PDFファイルを作成
    blankpage = source_path + 'blank.pdf'
    if not os.path.exists(blankpage):
        # os.remove(blankpage)
        page = canvas.Canvas(blankpage, pagesize=portrait(A4))
        # PDFファイルとして保存
        page.showPage()
        page.save()

    pdfs = []

    for f in filelist.index:
        file_pdf = tmpdir + (f.replace('.docx', '.pdf'))
        pdfs.append(file_pdf)
        if not args.skip_convert:
            print('Converting... ', f)        
            convert(docdir + f, file_pdf)
        
    out_file = str(source_path  +args.out_file)
    if os.path.exists(out_file): os.remove(out_file)

    pdfindex = pdf_merger(out_file, pdfs, args.add_blank)
    print("目次は以下です")
    # pdfindexのkeyでfilelistから目次名を取得
    for key, value in pdfindex.items():
        print(filelist.at[key.replace('.pdf', '.docx'), '目次'], ':', value)
    if args.wipe_tempdir : shutil.rmtree(tmpdir)

    print("ページ追加中")
    paged_file = str(source_path  + "paged.pdf")
    if os.path.exists(paged_file): os.remove(paged_file)
    add_page_number(out_file, paged_file, 0, args.from_pagenum)
    print("目次追加中")
    outlined_file = str(source_path  + "outlined.pdf")
    if os.path.exists(outlined_file): os.remove(outlined_file)
    add_outline(paged_file, outlined_file, pdfindex)
    if args.wipe_workingfiles:
        os.remove(out_file)
        os.remove(paged_file)
