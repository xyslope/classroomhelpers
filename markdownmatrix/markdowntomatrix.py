import os
import pandas as pd
import re
import numpy as np
from argparse import ArgumentParser

def get_options():
    argparser = ArgumentParser()
    argparser.add_argument('-p', '--source_path', type=str,
                           default = os.getcwd(),
                           help='Document root')
    argparser.add_argument('-s', '--source_file', type=str,
                           default = 'markdownsource.md',
                           help='Source Document')
    argparser.add_argument('-e', '--export_format', type=str,
                           default = 'html',
                           help='Export Format(txt/html/csv')
    argparser.add_argument('-f', '--file_format', type=str,
                           default = 'md',
                           help='Source Format(md/org')
    argparser.add_argument('-t', '--transpose', action='store_true',
                           default = False,
                           help='Transpose the results')
    return argparser.parse_args()


def export_html(df, f):
    htmldata='<html><body><table border="2" style="border-collapse: collapse; border-color: gray">'
    htmldata += '<tr>'
    htmldata += '<th scope="row"></th>'
    for col in df.columns:
        htmldata += '<th>{}</th>'.format(col.replace('\n', '<br />'))
    htmldata += '</tr> \n'
    for index, row in df.iterrows():
        htmldata += '<tr>'
        htmldata += '<th scope="row">{}</th>'.format(index)
        for item in row:
            if not item:
                item = '該当なし'
            htmldata += '<td>{}</td>'.format(item.replace('\n', '<br />'))
        htmldata += '</tr> \n'
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

def markdown_to_dataframe(src, maxlevel, header           ):
    cols = np.arange(maxlevel-1)
    df = pd.DataFrame(columns=cols)
    l1 = '^{} '.format(header)
    l2 = '^{} '.format(header+header)
    title = ''
    for l in src:
        result = re.match(l1, l)
        cell = re.match(l2, l)
        if result:
            title = l
            level = -1
        elif cell:
            level += 1
            body = re.sub(l2, '', l)
            df.loc[re.sub(l1, '', title), level] = body
        elif not title == '':
            if level == -1:
                title += '\n'+l
            else:
                body += '\n' + re.sub(l2, '', l)
                df.loc[re.sub(l1, '', title), level] = body
    # 本来、行のインデックスは別途作成したほうがいいけど、今はこんな感じで
    df.columns = df.iloc[0]
    return df.drop(df.index[[0]])

if __name__ == "__main__":
    args = get_options()
    source_path = args.source_path + '/'
    source_file = args.source_file
    with open (source_path+source_file, 'r', encoding='UTF-8') as f:    
        criteria = [s.strip() for s in f.readlines()]
    maxlevel = get_maxlevel(criteria)
    header = '\*' if args.file_format=='org' else '#'
    df = markdown_to_dataframe(criteria, maxlevel, header).fillna('該当なし')
    df = df.T if args.transpose else df

    if args.export_format == 'html':
        export_html(df, source_path + 
                os.path.basename(re.sub('md|org', 'html', source_file)))
    elif args.export_format == 'csv':
        export_csv(df, source_path + 
                os.path.basename(re.sub('md|org', 'csv', source_file)))
    else:
        export_txt(df)
