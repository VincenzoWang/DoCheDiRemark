import jieba
import xlrd
import json
import re
from collections import Counter
import matplotlib.pyplot as plt
#author：王天琛,这段代码主要用于最满意和最不满意评论的词频统计
#避免报错
jieba.setLogLevel(jieba.logging.INFO)
#解决matplotlib显示中文乱码的问题
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['font.family']='sans-serif'
#打开excel
worksheet = xlrd.open_workbook('./pinglun.xlsx')
sheet_names= worksheet.sheet_names()
# print(sheet_names)
sheet = worksheet.sheet_by_name(sheet_names[0])
rows = sheet.nrows # 获取行数
cols = sheet.ncols # 获取列数，尽管没用到
all_content = []

#最满意的
cols = sheet.col_values(0)[1:]
satisfied=' '.join(cols).replace(' ','')
#将无关紧要的词去掉
satisfied=re.sub('满意|可以|星越|非常|不错|地方|没有|这个|真的|比较','',satisfied)
#将全文分割，并将>=2的放入列表
xianni_words = [x for x in jieba.cut(satisfied) if len(x) >= 2]
c=Counter(xianni_words).most_common(10)
# print(json.dumps(c, ensure_ascii=False))
#画图
name_list = [x[0] for x in c]  # X轴的值
num_list = [x[1] for x in c]  # Y轴的值
b = plt.bar(range(len(num_list)), num_list, tick_label=name_list)  # 画图
plt.xlabel(u'词语')
plt.ylabel(u'次数')
plt.title(u'最满意的评论频率统计')
plt.show()  # 展示

#最不满意的
cols = sheet.col_values(1)[1:]
dissatisfied=' '.join(cols).replace(' ','')
#将无关紧要的词去掉
dissatisfied=re.sub('可以|就是|没有|这个|有点|满意|还是|感觉|时候|真的|问题|还有|不能|但是'
                    '|自动|每次|一个|关闭|知道|需要|希望|不是|自己|喜欢|地方|明显|吉利|出现|一样'
                    '|时间|不会|现在|而且|可能|一下|目前|星越|的话|个人|很多|有些|特别|什么|用车|'
                    '行车|设置|过程|缺点|一次|翠羽|只能|这么|一直|之前|如果|虽然|来说|行驶|只有|'
                    '容易|盲订|不过|竟然|影响|结果|不好|开启|体验|东西|毕竟|肯定|反应|偶尔|怎么|'
                    '已经|后面|一点|那么|支持|声音|高速|None|驾驶|提车|模式|比较','',dissatisfied)
#将全文分割，并将>=2的放入列表
xianni_words1 = [x for x in jieba.cut(dissatisfied) if len(x) >= 2]
c1=Counter(xianni_words1).most_common(10)
# print(json.dumps(c1, ensure_ascii=False))
#画图
name_list = [x[0] for x in c1]  # X轴的值
num_list1 = [x[1] for x in c1]  # Y轴的值
b1 = plt.bar(range(len(num_list1)), num_list1, tick_label=name_list)  # 画图
plt.xlabel(u'词语')
plt.ylabel(u'次数')
plt.title(u'最不满意的评论频率统计')
plt.show()  # 展示