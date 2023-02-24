import json
import os
import re
from collections import OrderedDict
import pandas as pd
import docx
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table, _Row
from docx.text.paragraph import Paragraph
from tqdm import tqdm

# df:pd.DataFrame = pd.read_excel(r'C:\Users\hujunchao\Desktop\库房500箱编号(2).xlsx', sheet_name='Sheet1')
# result_path = r'C:\Users\hujunchao\Desktop\库房500箱编号(2)_分組排序.xlsx'
#
# grouped = list(df.groupby('第三列'))
# grouped = [i[1].sort_values('老编号') for i in grouped]
# grouped = sorted(grouped, key=lambda x:x['老编号'].iloc[0])
#
# result = pd.concat(grouped, axis=0)
# result = result.reset_index(drop=True)
# result.to_excel(result_path, index=False)

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    elif isinstance(parent, _Row):
        parent_elm = parent._tr
    else:
        raise ValueError("something's not right")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def recognize_table(table: Table):
    result = []
    for row in table.rows:
        cells = row.cells
        result.append([i.text for i in cells])
    return result

def find_index(key, lst):
    for idx, item in enumerate(lst):
        if isinstance(item, str) and key.match(item) and isinstance(lst[idx+1], tuple):
            return idx
    return None

def find_key_table(lst: list):
    keys = [
        ('指数', re.compile('\s*指数$')),
        ('成交', re.compile('\s*成交$')),
        ('规模', re.compile('\s*规模$')),
        ('融资融券', re.compile('\s*融资融券$')),
        ('沪港通', re.compile('\s*沪港通$')),
        ('股票期权', re.compile('\s*股票期权$')),
        ('发行上市', re.compile('\s*发行上市$')),
        ('科创板项目动态', re.compile('\s*科创板项目动态$'))
    ]
    result = {}
    for key, pat in keys:
        key_table = []
        idx = find_index(pat, lst)
        if idx:
            for i in range(idx + 1, idx + 4):
                item = lst[i]
                if isinstance(item, tuple) and item[0] == 'table':
                    key_table.append(item[1])
                else:
                    break
            table = key_table[0]
            for t in key_table[1:]:
                table += t
            table = [' | '.join(i) for i in table]
            result[key] = table
    return result

def save_to_excel(dir, result, keys):
    for k in keys:
        if k in result:
            table = result[k]



def parse_docx():
    root = r'C:\Users\hujunchao\Documents\PdfDir\original\上交所docx'
    all_docx = os.listdir(root)
    result = []
    for doc_file in tqdm(all_docx):
        path = os.path.join(root, doc_file)
        doc = docx.Document(path)
        cur_doc_content = []
        for block in iter_block_items(doc):
            # read Paragraph
            if isinstance(block, Paragraph) and block.text:
                cur_doc_content.append(block.text)
            # read table
            elif isinstance(block, Table):
                table_content = recognize_table(block)
                cur_doc_content.append(('table', table_content))
        cur_doc_content = [i for i in cur_doc_content if i]
        tables = find_key_table(cur_doc_content)
        key = doc_file.replace('.docx', '.pdf')
        tables['file'] = key
        result.append(tables)
    with open('./上交所抽取结果.json', mode='w', encoding='utf8') as f:
        json.dump(result, f, indent=4, ensure_ascii=False)


if __name__ == '__main__':
    parse_docx()
