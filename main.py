import urllib.request
import re
import xlwt
import xlrd
from collections import Counter
import numpy as np
import matplotlib.pyplot as plt
plt.rcParams['font.family'] = ['sans-serif']
plt.rcParams['font.sans-serif'] = ['SimHei']


# 数据获得后进行正则匹配  获得所需数据：电影名称 分数 国家  年份  评价人数  其中排名按爬取顺序决定 不参与正则匹配
def write2excel(data):
    # 1. 设定正则匹配项
    pat1 = re.compile(r'"title":"(.*?)"')  # 名称
    pat2 = re.compile(r'"rating":\["(.*?)","\d+"\]')  # 分数
    part3 = re.compile(r'"regions":\["(.*?)"')  # 国家
    part4 = re.compile(r'"release_date":"(\d\d\d\d)')  # 年份
    part5 = re.compile(r'"vote_count":(.*?),')  # 评价人数
    # 2.获取关键数据
    data1 = pat1.findall(data, re.I)
    data2 = pat2.findall(data, re.I)
    data3 = part3.findall(data, re.I)
    data4 = part4.findall(data, re.I)
    data5 = part5.findall(data, re.I)
    # 数据导入Excel中 文件名为DouBanTop100Documentary.xls
    file = xlwt.Workbook('encoding = utf-8')
    sheet1 = file.add_sheet('sheet1', cell_overwrite_ok=True)

    sheet1.write(0, 0, "排名")
    sheet1.write(0, 1, "名称")
    sheet1.write(0, 2, "分数")
    sheet1.write(0, 3, "国家")
    sheet1.write(0, 4, "年份")
    sheet1.write(0, 5, "评价人数")
    for i in range(len(data1)):
        sheet1.write(i + 1, 0, i + 1)
        sheet1.write(i + 1, 1, data1[i])
        sheet1.write(i + 1, 2, data2[i])
        sheet1.write(i + 1, 3, data3[i])
        sheet1.write(i + 1, 4, data4[i])
        sheet1.write(i + 1, 5, data5[i])
    file.save('DouBanTop100Documentary.xls')


# 从Excel中读取数据为python数据类型
def writefromexcel(filename):

    data_excel = xlrd.open_workbook(filename)
    table = data_excel.sheets()[0]
    n_rows = table.nrows  # 获取该sheet中的有效行数
    Rank = table.col_values(colx=0, start_rowx=1, end_rowx=n_rows)
    Name = table.col_values(colx=1, start_rowx=1, end_rowx=n_rows)
    Score = table.col_values(colx=2, start_rowx=1, end_rowx=n_rows)
    Region = table.col_values(colx=3, start_rowx=1, end_rowx=n_rows)
    Year = table.col_values(colx=4, start_rowx=1, end_rowx=n_rows)
    Vote_count = table.col_values(colx=5, start_rowx=1, end_rowx=n_rows)
    return Rank, Name, Score, Region, Year, Vote_count


# 1. 前100名电影中，国家对应的数量 做饼图（出现多个国家的，可以保存一个就行）
def analysis1(region):
    count = dict(Counter(region))
    keys = count.keys()
    values = [0] * len(keys)
    i = 0
    for key in keys:
        values[i] = count[key]
        i += 1
    plt.title('各国占据Top100分布图')
    plt.pie(values, labels=keys)
    plt.show()


# 2. 每个电影的评价人数的柱状图
def analysis2(name, votecount):
    for i in range(len(votecount)):
        plt.bar(name[i], int(votecount[i]))
    plt.title("每部电影的评价人数的柱状图")
    # 设置x轴标签名
    plt.xlabel("名称")
    # 设置y轴标签名
    plt.ylabel("评价人数")
    # 显示
    plt.show()


# 3. 平均评价人数
def analysis3(votecount):
    votecount = list(map(int, votecount))
    print("100部电影的评价平均人数为：", np.mean(votecount))


# 4. 前100名中，各个国家的电影的平均评分对比柱状图
def analysis4(region, score):
    count = dict(Counter(region))
    keys = count.keys()
    keys_list = list(count.keys())
    values = [0] * len(keys)
    num = 0
    for i in region:
        index = keys_list.index(i)
        values[index] = (values[index] + float(score[num]))/2.0
        num += 1

    for i in range(len(keys)):
        plt.bar(keys_list[i], float(values[i]))
    plt.title("前100名中，各个国家的电影的平均评分对比柱状图")
    # 设置x轴标签名
    plt.xlabel("国家")
    # 设置y轴标签名
    plt.ylabel("平均分")
    # 显示
    plt.show()


# 5. 评分最高的10步电影的名称和评分的对比柱状图
def analysis5(Name, Score):
    for i in range(10):
        plt.bar(Name[i], float(Score[i]))
    plt.title("评分最高的10步电影的名称和评分的对比柱状图")
    # 设置x轴标签名
    plt.xlabel("名称")
    # 设置y轴标签名
    plt.ylabel("评分")
    # 显示
    plt.show()


#6.年份与对应的电影数量的折线图
def analysis6(year):
    count = dict(Counter(year))
    count = dict(sorted(count.items(), key=lambda x: x[0]))
    keys = count.keys()
    nums = [0] * len(keys)
    i = 0
    for key in keys:
        nums[i] = count[key]
        i += 1
    plt.title('年份与对应的电影数量的折线图')
    plt.plot(keys, nums, label="数目变化折线")
    plt.legend()  # 显示上面的label
    plt.xlabel('时间')
    plt.ylabel('数目')
    plt.show()

# 数据爬取
# headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:71.0) Gecko/20100101 Firefox/71.0"}  # 对抗网站的反爬取
# url = "https://movie.douban.com/j/chart/top_list?type=1&interval_id=100%3A90&action=&start=0&limit=100"  # 异步生成网页 从第一个开始：start=0， 爬取100份： limit=100
# res = urllib.request.Request(url, headers=headers)
# data = urllib.request.urlopen(res).read().decode()


if __name__ == '__main__':
    filename = 'DouBanTop100Documentary.xls'
    # write2excel(data)
    Rank, Name, Score, Region, Year, Vote_count = writefromexcel(filename)
    # 数据分析
    # 1. 前100名电影中，国家对应的数量 做饼图（出现多个国家的，可以保存一个就行）
    analysis1(Region)
    # 2. 每个电影的评价人数的柱状图
    analysis2(Name, Vote_count)
    # 3. 平均评价人数
    analysis3(Vote_count)
    # 4. 前100名中，各个国家的电影的平均评分对比柱状图
    analysis4(Region, Score)
    # 5.评分最高的10步电影的名称和评分的对比柱状图
    analysis5(Name, Score)
    # 6.年份与对应的电影数量的折线图
    analysis6(Year)