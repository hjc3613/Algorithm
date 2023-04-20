import os
from modification_1 import modification_1
import json
from tqdm import tqdm
directory = '最新文件抽取结果'
output_dir = '最新文件抽取结果-fix'

#在每个batch文件夹下建立子文件夹
# for foldername in os.listdir(directory):
#     folder_path = os.path.join(directory, foldername)
#     if os.path.isfile(folder_path):
#         continue
#     ok_folder_path = os.path.join(folder_path, "员工专业学历")
#     os.makedirs(ok_folder_path, exist_ok=True)


file_path = []
for root, dirs, files in os.walk(directory):
    if '__pycache__' in dirs:
        dirs.remove('__pycache__')
    # if '第六节业务与技术' in dirs:
    #     dirs.remove('第六节业务与技术')
    if '员工专业学历' in dirs:
        dirs.remove('员工专业学历')
    for filename in files:
        # if not any([x in filename for x in ['py','txt','json','zip']]):
        #     file_path.append(os.path.join(root, filename))
        if filename.endswith('.xlsx'):
            file_path.append(os.path.join(root, filename))

debug = {}
for i,f in tqdm(enumerate(file_path)):
    # print(f"Processing item {i + 1}/{len(file_path)}: {f}", end='\r')
    dir = os.path.dirname(f)
    file_name = os.path.basename(f)
    output_info = modification_1(f,f"{output_dir}/{os.path.split(dir)[-1]}/{file_name}")
    if output_info:
        debug[output_info[0]] = output_info[1]
print()
with open ('debug.json','w', encoding='utf8') as f:
    json.dump(debug,f, indent=4, ensure_ascii=False)