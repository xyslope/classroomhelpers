import os
import pandas as pd
import re
import numpy as np
from argparse import ArgumentParser

def get_options():
    argparser = ArgumentParser()
    argparser.add_argument('-b', '--add_blank', type=bool,
                           default=False,
                           help='If number of pages is odd, add a blank page.')
    argparser.add_argument('-p', '--source_path', type=str,
                           default = os.getcwd(),
                           help='Document root')
    argparser.add_argument('-s', '--source_file', type=str,
                           default = 'markdownsource.md',
                           help='Source Document')
    argparser.add_argument('-f', '--export_format', type=str,
                           default = 'html',
                           help='Export Format(txt/html/csv')
    return argparser.parse_args()


def export_html(df, f):
    header = True
    htmldata='<html><body><table>'
    for index, row in df.iterrows():
        htmldata += '<tr>'
        htmldata += '<th scope="row">{}</th>'.format(index)
        for item in row:
            if not item:
                item = '該当なし'
            if header:
                htmldata += '<th>{}</th>'.format(item.replace('\n', '<br />'))
            else:
                htmldata += '<td>{}</td>'.format(item.replace('\n', '<br />'))
        htmldata += '</tr> \n'
        header = False
    htmldata += '</table></body></html> \n'
    with open(f, encoding='UTF-8', mode = 'w') as f:
        f.write(htmldata)

def export_txt(df):
    print(df)

def export_csv(df, f):
    df.to_csv(f , encoding='UTF-8', header=True, index=True)

def get_maxlevel(src):
    # get max levels
    maxlevel = 0
    for l in src:
        if re.match('^# ', l):
            level = 0
        elif re.match('^## ', l):
            level += 1
            maxlevel = level if maxlevel < level else maxlevel
    return maxlevel

def markdown_to_dataframe(src, maxlevel):
    df = pd.DataFrame(index=[], columns=[np.arange(maxlevel-1)])
    for l in src:
        result = re.match('^# ', l)
        cell = re.match('^## ', l)
        if result:
            title = l
            level = -1
        elif cell:
            level += 1
            body = l.replace('## ', '')
            df.loc[title.replace('# ', ''), level] = body
        else:
            if level == -1:
                title += '\n'+l
            else:
                body += '\n' + l.replace('## ', '')
                df.loc[title.replace('# ', ''), level] = body
    return df

if __name__ == "__main__":
    args = get_options()
    source_path = args.source_path + '/'
    source_file = args.source_file
    with open ('freshpeople.md', 'r', encoding='UTF-8') as f:    
        criteria = [s.strip() for s in f.readlines()]
    maxlevel = get_maxlevel(criteria)
    df = markdown_to_dataframe(criteria, maxlevel).fillna('該当なし')
    if args.export_format == 'html':
        export_html(df, source_path + os.path.basename(source_file.replace('.md', '.html')))
    elif args.export_format == 'csv':
        export_csv(df, source_path + os.path.basename(source_file.replace('.md', '.csv')))
    else:
        export_txt(df)

