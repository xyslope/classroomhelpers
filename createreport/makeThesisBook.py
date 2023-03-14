import sys
import os
import shutil
import comtypes.client
import glob
import pathlib
import PyPDF2
from argparse import ArgumentParser
import pandas as pd

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
        if not key == 'ー':
            merger.add_outline_item(key, val, parent=None)
    merger.write(out_pdf)
    merger.close()
    return pdfindex


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

    blankpage = source_path + 'blank.pdf'
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
