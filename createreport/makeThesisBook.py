import sys
import os
import shutil
import comtypes.client
import pypdf
from pathlib import Path
from pypdf import PageObject
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
    argparser.add_argument('-b',
                           '--add_blank',
                           action="store_true",
                           help='If number of pages is odd, add a blank page.')
    argparser.add_argument(
        '-sc',
        '--skip_convert',
        help='Skip convert from docx to pdf.',
        action="store_true",
    )
    argparser.add_argument('-w',
                           '--wipe_tempdir',
                           action="store_true",
                           help='Wipe working directory.')
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
    argparser.add_argument('-o',
                           '--out_file',
                           type=str,
                           default='out.pdf',
                           help='Path to output directory')
    argparser.add_argument('-td',
                           '--temp_dir',
                           type=str,
                           default='temp',
                           help='Working directory')
    argparser.add_argument('-ww',
                           '--wipe_workingfiles',
                           action="store_true",
                           help='wipe working files')
    argparser.add_argument('-p',
                           '--from_pagenum',
                           type=int,
                           default=0,
                           help='Add page num from page x')
    return argparser.parse_args()


def convert(in_file, out_file):
    """convert word file to pdf
    in_file: word file with fullpath
    out_file: pdf file with fullpath
    """
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(str(in_file))
    doc.SaveAs(str(out_file), FileFormat=17)
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
    merger = pypdf.PdfWriter()
    for pdf in pdfs:
        pdftitle = filelist.at[os.path.basename(pdf).replace('.pdf', '.docx'),
                               '目次']
        pdfindex[pdftitle] = len(merger.pages) + 1
        merger.append(pdf, pdftitle)
        if add_blank & (len(merger.pages) % 2) == 1:
            merger.append(blankpage)
        print(os.path.basename(pdf), '(pages: ', len(merger.pages), ')')
    merger.write(out_pdf)
    return pdfindex


def add_page_number(input_file: str,
                    output_file: str,
                    start_num: int = 1,
                    record_from: int = 0):
    """
    既存PDFにページ番号を追加する
    """
    # 既存PDF（ページを付けるPDF）
    fi = open(input_file, 'rb')
    pdf_reader = pypdf.PdfReader(fi)
    pages_num = len(pdf_reader.pages)

    # ページ番号を付けたPDFの書き込み用
    pdf_writer = pypdf.PdfWriter()

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

    # ページ番号だけのPDFをメモリから読み込み（seek操作はpypdfに実装されているので不要）
    pdf_num_reader = pypdf.PdfReader(bs)

    # 既存PDFに１ページずつページ番号を付ける
    for i in range(0, pages_num):
        # 既存PDF
        pdf_page = pdf_reader.pages[i]
        # ページ番号だけのPDF
        pdf_num = pdf_num_reader.pages[i]
        if i >= record_from - 1:
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
        c.drawRightString(page_size[0] - 25, PAGE_BOTTOM,
                          "-" + str(page_num) + "-")
    else:
        c.drawString(25, PAGE_BOTTOM, "-" + str(page_num) + "-")
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
    reader = pypdf.PdfReader(infile)
    writer = pypdf.PdfWriter()
    writer.append_pages_from_reader(reader)
    for key, val in pdfindex.items():
        if not key == '-':
            print(key, val)
            writer.add_outline_item(key, val - 1, parent=None)  # add bookmark
    with open(outfile, "wb") as fp:
        writer.write(fp)

def setfile(source, file_name):
    file_path = Path(source) / file_name.lstrip("\\/")
    file_path.unlink(missing_ok=True)
    return file_path

if __name__ == "__main__":
    args = get_options()
    source_path = args.source_path
    docdir = Path(source_path) / args.doc_dir.lstrip("\\/")
    if not docdir.exists():
        print('ソースディレクトリが存在しません。')
        sys.exit()
    tmpdir = Path(source_path) / args.temp_dir.lstrip("\\/")
    if not args.skip_convert:
        if tmpdir.exists(): shutil.rmtree(tmpdir)
        tmpdir.mkdir()
    out_file = setfile(source_path, args.out_file)
    paged_file = setfile(source_path, "paged.pdf")
    outlined_file = setfile(source_path,   "outlined.pdf")


    # A4の新規PDFファイルを作成
    blankpage = Path(source_path) / 'blank.pdf'
    if not blankpage.exists():
        # os.remove(blankpage)
        page = canvas.Canvas(blankpage, pagesize=portrait(A4))
        # PDFファイルとして保存
        page.showPage()
        page.save()

    pdfs = []

    filelist = pd.read_csv(Path(source_path) / 'filelist.csv')
    filelist = filelist.set_index('file')
    for f in filelist.index:
        file_pdf = Path(tmpdir) / f.replace('.docx', '.pdf')
        pdfs.append(file_pdf)
        if not args.skip_convert:
            print('Converting... ', f)
            convert(Path(docdir) / f, file_pdf)

    pdfindex = pdf_merger(out_file, pdfs, args.add_blank)
    print("目次は以下です")
    # pdfindexのkeyでfilelistから目次名を取得
    for key, value in pdfindex.items():
        print(key, ':', value)
#        print(filelist.at[key.replace('.pdf', '.docx'), '目次'], ':', value)

    print("ページ追加中")
    add_page_number(out_file, paged_file, 1, args.from_pagenum)
    print("目次追加中")
    add_outline(paged_file, outlined_file, pdfindex)

    if args.wipe_workingfiles:
        out_file.unlink()
        paged_file.unlink()
    if args.wipe_tempdir: shutil.rmtree(tmpdir)
