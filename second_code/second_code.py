# -*- coding: utf-8 -*-
"""
功能：提取市净率值
@author: zhw
@time:2016-2-11
"""
#获取数据
import xlrd,xlwt
data = xlrd.open_workbook('./data_sources/F100401A.xls').sheet_by_index(0)
x1 = data.col_values(0)
x2 = data.col_values(1)
x3 = data.col_values(3)

stocks = x1[1:]
times = x2[1:]
f = x3[1:]


#【定义】股票种类
stock_names = []

#【定义】数据中每一类股票占多少条数目，每只股票索引开始值
stock_total_indexs = []

stock_names.append(stocks[0])
stock_total_indexs.append(0)


#提取股票种类和每类股票起始索引值
k = 0

for stock in stocks :
    if stock_names[k] != stock:
        stock_total_indexs.append(stocks.index(stock))
        k = k + 1;
        stock_names.append(stock)


#【函数】提取天数数据的年月
def get_year_month(time):
    temp_time = time.split('-')
    month = temp_time[1]
    year = temp_time[0]
    return int(year)*100+int(month)
assert get_year_month('2011-02-06') == 201102

#查看最多丢多少季度数据，发现最大丢了16-1 = 15个，分类好麻烦，只好暴力提取了-_-！
min_value = 16
for k in xrange(len(stock_names)-1):
    a = stock_total_indexs[k+1]-stock_total_indexs[k]
    if min_value > a:
        min_value = a
        print a


#提取每只股票16个季度的净率值，无数据则赋值为 1
#【定义】标记遍历过的索引值总数，方便从原始数据中取数
sum_index = 0
#【定义】二重列表，用于存放每只股票每个季度净率值
ffff = []

for stock_index in xrange(len(stock_names)):
    #【定义】每只股票对应输出的净率值列表
    f_each = [1]*16
    if stock_index < len(stock_names)-1:
        temp_quarter_16 = stock_total_indexs[stock_index+1]-stock_total_indexs[stock_index]
        for quarter_index in xrange(temp_quarter_16):
            temp_time_value = times[sum_index+quarter_index]
            if get_year_month(temp_time_value) == 201103:
                f_each[0] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201106:
                f_each[1] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201109:
                f_each[2] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201112:
                f_each[3] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201203:
                f_each[4] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201206:
                f_each[5] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201209:
                f_each[6] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201212:
                f_each[7] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201303:
                f_each[8] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201306:
                f_each[9] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201309:
                f_each[10] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201312:
                f_each[11] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201403:
                f_each[12] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201406:
                f_each[13] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201409:
                f_each[14] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201412:
                f_each[15] = f[sum_index+quarter_index]
        sum_index = sum_index + temp_quarter_16
        ffff.append(f_each)
    else:
        temp_quarter_16 = len(f)-stock_total_indexs[stock_index]
        for quarter_index in xrange(temp_quarter_16):
            temp_time_value = times[sum_index+quarter_index]
            if get_year_month(temp_time_value) == 201103:
                f_each[0] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201106:
                f_each[1] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201109:
                f_each[2] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201112:
                f_each[3] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201203:
                f_each[4] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201206:
                f_each[5] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201209:
                f_each[6] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201212:
                f_each[7] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201303:
                f_each[8] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201306:
                f_each[9] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201309:
                f_each[10] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201312:
                f_each[11] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201403:
                f_each[12] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201406:
                f_each[13] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201409:
                f_each[14] = f[sum_index+quarter_index]
            elif get_year_month(temp_time_value) == 201412:
                f_each[15] = f[sum_index+quarter_index]
        sum_index = sum_index + temp_quarter_16
        ffff.append(f_each)

#结果写入excel
book = xlwt.Workbook(encoding='utf-8',style_compression=0)
file_xls = book.add_sheet('f',cell_overwrite_ok = True) 
file_xls.write(0,0,'stock types')
file_xls.write(0,1,'f')
for s_index in xrange(len(ffff)):
    file_xls.write(s_index+1,0,stock_names[s_index])
    for si in xrange(len(ffff[s_index])):
        file_xls.write(s_index+1,si+1,ffff[s_index][si])

book.save('./results/result.xls')












