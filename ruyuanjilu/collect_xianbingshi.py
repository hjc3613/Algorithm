import json
from pymongo import MongoClient

def collect_xianbingshi():
    client = MongoClient('localhost:27017')
    db = client.get_database('jhdata')
    collection = db.get_collection('ruyuanjilu')
    xianbingshi = set()
    for item in collection.find():
        xianbingshi_str = item['ruyuanjilu'].get('history_of_present_illness', {}).get('src')
        if xianbingshi_str:
            xianbingshi.add(xianbingshi_str)
    
    db = client.get_database('data')
    collection = db.get_collection('ruyuanjilu')
    xianbingshi = set()
    for item in collection.find():
        xianbingshi_str = item['ruyuanjilu'].get('history_of_present_illness', {}).get('src')
        if xianbingshi_str:
            xianbingshi.add(xianbingshi_str)

    with open('data/xianbingshi_total.txt', mode='w', encoding='utf8') as f:
        f.write('\n'.join(xianbingshi))

def convert_to_biaozhu_format():
    file = 'data/xianbingshi.txt'
    with open(file, encoding='utf8') as f:
        lines = [i.strip() for i in f.readlines()]
    
    for idx, line in enumerate(lines):
        with open(f'待标注/{str(idx).zfill(5)}.txt', mode='w', encoding='utf8') as f:
            f.write(line)

if __name__ == '__main__':
    # collect_xianbingshi()
    convert_to_biaozhu_format()