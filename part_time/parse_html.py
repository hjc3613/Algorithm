import json
import re
import os
import time
from typing import List
import shutil

import pandas as pd
from tqdm import tqdm

from bs4 import BeautifulSoup
from flatten_json import flatten

def find_all_titles(html:List[BeautifulSoup]):
    titles = []
    for idx, i in enumerate(html):
        for j in i.find_all('span'):
            if j.attrs.get('style')=='font-weight:bold;' and len(i.text) > 4:
                titles.append((idx, i.text))
                break
    return titles

def is_hierarchy_1(text):
    patterns = [
        (re.compile(r'管理层讨论与分析'), '第三节 管理层讨论与分析')
    ]
    for p, key in patterns:
        if p.search(text):
            return key
    return False

def is_hierarchy_2(text):
    text = re.sub(r'\s+', '', text)
    patterns = [
        (re.compile(r'[一二三四五六七八九].{0,2}经营情况讨论与分析'), '经营情况讨论与分析'),
        (re.compile(r'[一二三四五六七八九].{0,2}(报告期内公司所从事的主要业务、经营模式、行业情况及研发情况说明|报告期内公司所属行业及主营业务情况说明)'), '报告期内公司所从事的主要业务、经营模式、行业情况及研发情况说明'),
        (re.compile(r'[一二三四五六七八九].{0,2}报告期内核心竞争力分析'), '报告期内核心竞争力分析'),
        (re.compile(r'[一二三四五六七八九].{0,2}风险因素'), '风险因素'),
        (re.compile(r'[一二三四五六七八九].{0,2}报告期内主要经营情况'), '报告期内主要经营情况'),
        (re.compile(r'[一二三四五六七八九].{0,4}关于公司未来发展的讨论与分析'), '公司关于公司未来发展的讨论与分析'),
        (re.compile(r'[一二三四五六七八九].{0,2}公司因不适用准则规定或国家秘密'), '公司因不适用准则规定或国家私密、未按准则披露的情况和原因说明'),
    ]
    for p, key in patterns:
        if p.search(text):
            return key
    return False

def is_hierarchy_3(text):
    text = re.sub(r'\s+', '', text)
    if '三、财务风险' in text:
        a = 1
    patterns = [
        # 二 报告期内公司所从事的主要业务、经营模式、行业情况及研发情况说明
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}主要业务、主要产品或服务情况'), '主要业务、主要产品或服务情况'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}主要经营模式'), '主要经营模式'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}所处行业情况'), '所处行业情况'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}核心技术与研发进展'), '核心技术与研发进展'),
        # 三 报告期内核心竞争力分析
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}核心竞争力分析'), '核心竞争力分析'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}报告期内发生的导致公司核心竞争力受到严重影响的事件、影响分析及应对措施'), '报告期内发生的导致公司核心竞争力受到严重影响的事件、影响分析及应对措施'),
        # 四 风险因素
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}尚未盈利的风险'), '尚未盈利的风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}业绩大幅下滑或亏损的风险'), '业绩大幅下滑或亏损的风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}核心竞争力风险'), '核心竞争力风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}经营风险'), '经营风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}财务风险'), '财务风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}行业风险'), '行业风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}宏观环境风险'), '宏观环境风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}存托凭证相关风险'), '存托凭证相关风险'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}其他重大风险'), '其他重大风险'),
        # 五 报告期内主要经营情况
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}主营业务分析'), '主营业务分析'),
        # 六 关于公司未来发展的讨论与分析.{0,3}
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}行业格局和趋势'), '行业格局和趋势'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}公司发展战略'), '公司发展战略'),
        (re.compile(r'[(（]?[一二三四五六七八九][）)]?\s*.{0,3}经营计划'), '经营计划'),
        # (re.compile(r'[(（][一二三四五六七八九][）)]\s*其他'), '其他'),
    ]
    for p, key in patterns:
        if p.search(text):
            return key
    return False

def is_hierarchy_4(text):
    text = re.sub(r'\s+', '', text)
    patterns = [
        (re.compile(r'[1-9].{0,3}行业的发展阶段、基本特点、主要技术门槛'), '行业的发展阶段、基本特点、主要技术门槛'),
        (re.compile(r'[1-9].{0,3}公司所处的行业地位分析及其变化情况'), '公司所处的行业地位分析及其变化情况'),
        (re.compile(r'[1-9].{0,3}报告期内新技术、新产业、新业态、新模式的发展情况和未来发展趋势'), '报告期内新技术、新产业、新业态、新模式的发展情况和未来发展趋势'),
        (re.compile(r'[1-9].{0,3}核心技术及其先进性以及报告期内的变化情况'), '核心技术及其先进性以及报告期内的变化情况'),
        (re.compile(r'[1-9].{0,3}报告期内获得的研发成果'), '报告期内获得的研发成果'),
        (re.compile(r'[1-9].{0,3}研发投入情况表'), '研发投入情况表'),
        (re.compile(r'[1-9].{0,3}在研项目情况'), '在研项目情况'),
        (re.compile(r'[1-9].{0,3}研发人员情况'), '研发人员情况'),
        # (re.compile(r'[1-9].{0,3}其他说明'), '其他说明'),
        (re.compile(r'[1-9].{0,3}技术迭代风险'), '技术迭代风险'),
        (re.compile(r'[1-9].{0,3}研发失败风险'), '研发失败风险'),
        (re.compile(r'[1-9].{0,3}技术未能形成产品或实现产业化风险'), '技术未能形成产品或实现产业化风险'),
        (re.compile(r'[1-9].{0,3}依赖核心技术人员的风险'), '依赖核心技术人员的风险'),
        (re.compile(r'[1-9].{0,3}市场风险'), '市场风险'),
        (re.compile(r'[1-9].{0,3}丧失主要经营资质的风险'), '丧失主要经营资质的风险'),
        (re.compile(r'[1-9].{0,3}规模扩大导致的经营管理风险'), '规模扩大导致的经营管理风险'),
        (re.compile(r'[1-9].{0,3}毛利率下滑的风险'), '毛利率下滑的风险'),
        (re.compile(r'[1-9].{0,3}固定成本上升的风险'), '固定成本上升的风险'),
    ]
    for p, key in patterns:
        if p.search(text):
            return key
    return False

def find_main_titles(all_titles):
    hierarchy_1 = []
    hierarchy_2 = []
    hierarchy_3 = []
    hierarchy_4 = []
    for lineno, title in all_titles:
        hierarchy_1_key = is_hierarchy_1(title)
        hierarchy_2_key = is_hierarchy_2(title)
        hierarchy_3_key = is_hierarchy_3(title)
        hierarchy_4_key = is_hierarchy_4(title)
        if hierarchy_1_key:
            hierarchy_1.append((lineno,hierarchy_1_key))
        if hierarchy_2_key:
            hierarchy_2.append((lineno, hierarchy_2_key))
        if hierarchy_3_key:
            hierarchy_3.append((lineno, hierarchy_3_key))
        if hierarchy_4_key:
            hierarchy_4.append((lineno, hierarchy_4_key))
    return  hierarchy_1, hierarchy_2, hierarchy_3,hierarchy_4

def extract_3_chapter(html, filename):
    titles = [(idx,i) for idx,i in enumerate(html) if
              re.search('style="font-weight:bold;"', i)
              and (re.search(r'第[二三四]节\s*(</?span.*?>)*管理层讨论与分析', i) or re.search(r'第[三四五]节\s*(</?span.*?>)*(公司治理|董事会报告)',i) )
              and re.search(r'<h\d>', i)
              ]
    if not len(titles) == 2:
        chapter_34 = [i for i in html if re.search('style="font-weight:bold;"', i) and re.search(r'<h\d>',i) and re.search(r'第.{1,2}节',i)]
        chapter_34 = [BeautifulSoup(i).text.strip() for i in chapter_34]
        print(f'{filename} 没有找到章节名字，疑似章节名：{" | ".join(chapter_34)}')
        return None
    start, end = titles[0][0], titles[1][0]
    return html[start:end]

def create_col(heirarchy, length):
    result = [None]*length
    for lineno, key in heirarchy:
        result[lineno] = key
    return result

def generate_to_json(df:pd.DataFrame, cur_depth, max_depth=4, col_name='hierarchy'):
    df.reset_index(inplace=True, drop=True)
    cur_col = col_name + str(cur_depth)
    # 达到最大深度，或者当前分组中没有title，结束
    # if cur_depth > max_depth or not any(list(df[cur_col])):
    if cur_depth > max_depth:
        return '\n'.join(df['text'])
    if pd.isna(df[cur_col][0]) or pd.isnull(df[cur_col][0]):
        df[cur_col][0] = '未分组部分'
    df[cur_col] = df[cur_col].fillna(method='ffill')
    result = {}
    for key, sub_df in df.groupby(cur_col):
        val = generate_to_json(sub_df, cur_depth+1, max_depth=max_depth, col_name=col_name)
        result[key] = val
    # if '未分组部分' in result:
    #     result.pop('未分组部分')
    return result


def main():
    root = r'E:\part_time\管理层讨论与分析\PDF2HTML\2022'
    save_dir = 'structured_file3'
    os.makedirs(save_dir, exist_ok=True)
    all_htmls = os.listdir(root)
    # all_htmls = ['688161_20220330_2_j4YGw1bm.htm']
    all_htmls = [i for i in all_htmls if i.endswith('.htm')]
    for html_file in tqdm(all_htmls):
        with open(os.path.join(root, html_file), encoding='utf8') as f:
            lines = f.read()
            lines = lines.replace('<br><br>', '\n').replace('&nbsp;', ' ')
            lines = lines.split('\n')
        chapter_3 = extract_3_chapter(lines, filename=html_file)
        if chapter_3 is None:continue
        chapter_3 = [BeautifulSoup(i, 'lxml') for i in chapter_3]
        # 所以黑体字段落
        titles = find_all_titles(chapter_3)
        # pd.DataFrame.from_records(titles, columns=['lineno', 'title']).to_csv(html_file.replace('.htm', '.csv'))
        # 所有需要的标题
        hierarchy_1, hierarchy_2, hierarchy_3,hierarchy_4 = main_titles = find_main_titles(titles)
        # 构造 DataFrame
        chapter_3_text = [(i.text or '') for i in chapter_3]
        chapter_3_hierarchy_1 = create_col(hierarchy_1, len(chapter_3_text))
        chapter_3_hierarchy_2 = create_col(hierarchy_2, len(chapter_3_text))
        chapter_3_hierarchy_3 = create_col(hierarchy_3, len(chapter_3_text))
        chapter_3_hierarchy_4 = create_col(hierarchy_4, len(chapter_3_text))
        df = pd.DataFrame.from_dict({
            'text':chapter_3_text,
            'hierarchy1':chapter_3_hierarchy_1,
            'hierarchy2':chapter_3_hierarchy_2,
            'hierarchy3':chapter_3_hierarchy_3,
            'hierarchy4':chapter_3_hierarchy_4
        })
        result = generate_to_json(df, cur_depth=1, max_depth=4, col_name='hierarchy')
        save_name = html_file.replace('.htm', '.json')
        with open(os.path.join(save_dir, save_name), mode='w', encoding='utf8') as f:
            json.dump(result, f, indent=4, ensure_ascii=False)
        df.to_excel(os.path.join(save_dir, html_file.replace('.htm', '.xlsx')))
    return

def json2excel():
    json_files = os.listdir('structured_file')
    json_files = [i for i in json_files if i.endswith('.json')]
    all_json = []
    for json_file in json_files:
        with open(os.path.join('structured_file', json_file), encoding='utf8') as f:
            data = json.load(f)
            data['fileName'] = json_file
            all_json.append(flatten(data))
    df = pd.DataFrame.from_records(all_json)
    df = df.sort_index(axis=1)
    new_df = {}
    for col in df.columns:
        col_path = col.split('_')
        col_path = col_path + (4-len(col_path))*['']
        new_df[col] = col_path + list(df[col])
    new_df = pd.DataFrame.from_dict(new_df)
    new_header = new_df.iloc[0]  # grab the first row for the header
    new_df = new_df[1:]  # take the data less the header row
    new_df.columns = new_header  # set the header row as the df header
    new_df.to_excel('all_json.xlsx')

def get_new_files():
    new_dir = r'W:\datasets\2021年科创板年报(1)'
    old_dir = r'W:\datasets\2021年科创板年报'
    new_files = set(os.listdir(new_dir))
    old_files = set(os.listdir(old_dir))
    new_files = new_files - old_files
    src_files = [os.path.join(new_dir, i) for i in new_files]
    tgt_files = [os.path.join(r'C:\Users\hujunchao\Documents\PdfDir\original\2021年科创板年报', i) for i in new_files]
    for src, tgt in zip(src_files, tgt_files):
        shutil.copy(src, tgt)

if __name__ == '__main__':
    # 第一步：解析html，生成结构化的json，保存到structured_file下
    main()
    # 第二步：读取structured_file下的json，生成Excel文件
    # json2excel()
    # get_new_files()