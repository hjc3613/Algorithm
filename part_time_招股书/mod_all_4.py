import os
from modification_4 import modification_4
import json
from tqdm import tqdm
directory = '最新文件抽取结果-fix3'
output_dir = '最新文件抽取结果-fix4'
os.makedirs(output_dir, exist_ok=True)


#在每个batch文件夹下建立子文件夹(只需要执行一次)
'''
for foldername in os.listdir(directory):
    folder_path = os.path.join(directory, foldername)
    ok_folder_path = os.path.join(folder_path, "经营成果分析")
    os.makedirs(ok_folder_path, exist_ok=True)
'''
ignore = ['__pycache__','no_dup','主营业务收入，按产品、地域等划分',
          '员工专业学历','第六节业务与技术','前五大客户销售','经营成果分析']
file_path = []
for root, dirs, files in os.walk(directory):
    for ig in ignore:
        if ig in dirs:
            dirs.remove(ig)
    for filename in files:
        # if not any([x in filename for x in ['py','txt','json','zip']]):
        #     file_path.append(os.path.join(root, filename))
        if filename.endswith('.xlsx'):
            file_path.append(os.path.join(root, filename))

debug = {}
for i,f in tqdm(enumerate(file_path)):
    print(f"Processing item {i + 1}/{len(file_path)}: {f}", end='\r')
    dir = os.path.dirname(f)
    file_name = os.path.basename(f)
    output = modification_4(f, f"{output_dir}/{os.path.split(dir)[-1]}/{file_name}")
    if output:
        debug[output[0]] = output[1]
print()
with open ('debug_经营成果分析.json','w') as f:
    json.dump(debug,f,indent=4)