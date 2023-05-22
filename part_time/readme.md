### 抽取思路：

1. #### 解析PDF，考虑到要通过黑体字识别章节、目录，经测试发现 ABBYY识别效果较好，且能保证黑体字信息几乎无丢失，所以第一步用ABBYY将PDF解析为html文件

2. #### 处理html文件，首先是找到 **管理层讨论与分析 章节**，大部分是在第三章，少数不是，当前代码还未优化此处，通过以下逻辑定位该章节：

   ```python
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
   ```

3. #### 定位出 **管理层讨论与分析** 章节后，该查找结构化信息，结构化字段主要包括：

   一、报告期内公司所从事的主要业务、经营模式、行业情况及研发情况说明

   ​	1、主要业务、主要产品或服务情况

   ​	2、主要经营模式

   ​	3、所处行业情况

   ​	4、核心技术与研发进展

   二、经营情况讨论与分析

   三、报告期内核心竞争力分析

   ​	1、核心竞争力分析

   ​	2、报告期内发生的导致公司核心竞争力受到严重影响的事件、影响分析及应对措施

   四、风险因素

   ​	1、尚未盈利的风险

   ​	2、业绩大幅下滑或亏损的风险

   ​	3、核心竞争力风险

   ​	4、经营风险

   ​	5、财务风险

   ​	6、行业风险

   ​	7、宏观环境风险

   ​	8、存托凭证相关风险

   ​	9、其他重大风险

   五、报告期内主要经营情况

   ​	1、主营业务分析

   六、公司关于公司未来发展的讨论与分析

   ​	1、行业格局和趋势

   ​	2、公司发展战略

   ​	3、经营计划

   七、公司因不适用准则规定或国家私密、未按准则披露的情况和原因说明

   

   注：*以上结果是基于2021年的数据总结出的结构，不一定完善，可以适当增删改查。*

   为了得到以上结构，采用了以下三步：

   1) 先从第三章的文本列表中得到所有的黑体字，作为章节抽取的候选项：

   ```python
   # 所以黑体字段落
   titles = find_all_titles(chapter_3)
   def find_all_titles(html:List[BeautifulSoup]):
       titles = []
       for idx, i in enumerate(html):
           for j in i.find_all('span'):
               if j.attrs.get('style')=='font-weight:bold;' and len(i.text) > 4:
                   titles.append((idx, i.text))
                   break
       return titles
   ```

   2) 再从候选项中找到通过预设规则匹配到的章节名：

      ```python
      hierarchy_1, hierarchy_2, hierarchy_3,hierarchy_4 = main_titles = find_main_titles(titles)
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
      ```

3) 构造dataframe：

   | text                                                         | hierarchy1              | hierarchy2                                                   | hierarchy3                   | hierarchy4 |
   | ------------------------------------------------------------ | ----------------------- | ------------------------------------------------------------ | ---------------------------- | ---------- |
   | 第三节 管理层讨论与分析                                      | 第三节 管理层讨论与分析 |                                                              |                              |            |
   | 一、报告期内公司所属行业及主营业务情况说明                   | 第三节 管理层讨论与分析 | 报告期内公司所从事的主要业务、经营模式、行业情况及研发情况说明 |                              |            |
   | （一） 主要业务、主要产品或服务情况                          | 第三节 管理层讨论与分析 |                                                              | 主要业务、主要产品或服务情况 |            |
   | 公司是行业领先的工业自动化测试设备与整线系统解决方案提供商。基于在电子、光学、声 学、射频、机器视觉、机械自动化等多学科交叉融合的核心技术为客户提供从整机、系统、模块、 SIP、芯片各个工艺节点的自动化测试设备。公司产品主要应用于LCD与OLED平板显示及微显示、 半导体集成电路、消费电子可穿戴设备、新能源汽车等行业。作为一家专注于全球化专业检测领 域的高科技企业，公司坚持在技术研发、产品质量、技术服务上为客户提供具有竞争力的解决方 案，在各类数字、模拟、射频等高速、高频、高精度信号板卡、基于平板显示检测的机器视觉图 像算法，以及配套各类高精度自动化与精密连接组件的设计制造能力等方面具备较强的竞争优势 和自主创新能力。 | 第三节 管理层讨论与分析 |                                                              |                              |            |
   | 报告期内公司主要产品情况见下表:                              | 第三节 管理层讨论与分析 |                                                              |                              |            |
   | 产品 产品                                                    | 第三节 管理层讨论与分析 |                                                              |                              |            |
   |                                                              | 第三节 管理层讨论与分析 |                                                              |                              |            |

4) 以上dataframe其实已经是可用的结构化数据，为方便下游处理，可以转换成json。也可以直接使用。转成json时，会根据hierarchy*的值做聚合。具体逻辑见代码。逻辑不够清晰，可以随意优化