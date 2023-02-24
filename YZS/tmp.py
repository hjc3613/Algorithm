import pandas as pd

def add_no_injure_col():
    map_file = r'C:\Users\hujunchao\Desktop\YZS\第二批对比测试集_全部.csv'
    map_df = pd.read_csv(map_file)
    injure_no_injure_map = {}
    for tp in map_df[['诊断名字及代码', '所有伤', '所有非伤']].itertuples(index=False):
        injure_no_injure_map[tp[0]] = (tp[1], tp[2])
    injure_no_injure_map

    to_biaozhu_file = r'C:\Users\hujunchao\Desktop\YZS\对比测试集_抽样fraq_0.5.csv'
    to_biaozhu_df = pd.read_csv(to_biaozhu_file)
    no_injure = []
    for main_diag in to_biaozhu_df['诊断名字及代码']:
        assert main_diag in injure_no_injure_map
        no_injure.append(injure_no_injure_map[main_diag][1])
    to_biaozhu_df['所有非伤'] = no_injure
    to_biaozhu_df.to_csv(r'C:\Users\hujunchao\Desktop\YZS\对比测试集_抽样fraq_0.5_add_noInjure.csv')

if __name__ == '__main__':
    add_no_injure_col()