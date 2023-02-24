import json
import os
from collections import Counter
def analize_zhuyingyewu_keys():
    '''
    分析 主营业务 检测到的key
    :return:
    '''
    counter = Counter()
    with open('主营业务分析.txt', encoding='utf8') as f:
        lines = [i.strip() for i in f.readlines() if i.strip() and i.startswith('zg')]
        lines_shorter = [i for i in lines if len(i.split(' ', maxsplit=2)[2]) < 30]
        lines = [i.split(' ', maxsplit=2)[2] for i in lines]
        lines = [i for i in lines if len(i) < 30]
        counter.update(lines)
    counter = counter.most_common()
    print('total count: ', sum([count for k, count in counter]))
    counter = [k for k,count in counter]
    with open('营业收入关键词结果分析.json', encoding='utf8', mode='w') as f:
        json.dump(counter, f, ensure_ascii=False, indent=4)
    with open('主营业务分析-shorter.txt', encoding='utf8', mode='w') as f:
        f.write('\n'.join(lines_shorter))
    return

def check_guben_bianhua():
    root = '招股书extracted3'
    for file in os.listdir(root):
        with open(os.path.join(root, file), encoding='utf8') as f:
            data = json.load(f)
        guben_bianhua = data['发行前后股本变化情况']
        if len(guben_bianhua) == 0:
            print(file)

if __name__ == '__main__':
    # analize_zhuyingyewu_keys()
    check_guben_bianhua()