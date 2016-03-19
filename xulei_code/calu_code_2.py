# -*- coding: utf-8 -*-
"""
about how to caculate the skewness about stocks 
@author: zhw
@time:2016-2-22
@version:v5

"""
#获取数据
import xlrd,xlwt
data = xlrd.open_workbook('./data_sources/TRD_Dalyr.xls').sheet_by_index(0)
x1 = data.col_values(0)
x2 = data.col_values(1)
x3 = data.col_values(3)
x4 = data.col_values(2)

stocks = x1[1:]
times = x2[1:]
rit = x3[1:]
marktv = x4[1:]

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

#【定义】每只股票16个季度每个季度天数,二维列表，[0][2]->表示第一只股票第三季度天数
time_16_days = []

#【函数1】提取天数数据的月份
def get_month(time):
    temp_time = time.split('-')
    month = temp_time[1]
    return int(month)
assert get_month('2011-02-06') == 02


#【函数2】提取天数数据的年月
def get_year_month(time):
    temp_time = time.split('-')
    month = temp_time[1]
    year = temp_time[0]
    return int(year)*100+int(month)
assert get_year_month('2011-02-06') == 201102




#【定义】季度索引量，总16个
quarter_index = 0

#方法一：提取每只股票每个季度天数
#【定义】标记遍历过的索引值总数，方便从原始数据中取数
sum_index = 0

for stock_index in xrange(len(stock_names)):
    #【定义】每个季度天数临时变量，最大为16季度
    quarter_days = [0]*16
    if stock_index < len(stock_names)-1:
        temp_quarter_16 = stock_total_indexs[stock_index+1]-stock_total_indexs[stock_index]
        for quarter_index in xrange(temp_quarter_16):
            temp_time_value = times[sum_index+quarter_index]
            if get_year_month(temp_time_value) <= 201103:
                quarter_days[0] += 1
            elif get_year_month(temp_time_value) <= 201106:
                quarter_days[1] += 1
            elif get_year_month(temp_time_value) <= 201109:
                quarter_days[2] += 1
            elif get_year_month(temp_time_value) <= 201112:
                quarter_days[3] += 1
            elif get_year_month(temp_time_value) <= 201203:
                quarter_days[4] += 1
            elif get_year_month(temp_time_value) <= 201206:
                quarter_days[5] += 1
            elif get_year_month(temp_time_value) <= 201209:
                quarter_days[6] += 1
            elif get_year_month(temp_time_value) <= 201212:
                quarter_days[7] += 1
            elif get_year_month(temp_time_value) <= 201303:
                quarter_days[8] += 1
            elif get_year_month(temp_time_value) <= 201306:
                quarter_days[9] += 1
            elif get_year_month(temp_time_value) <= 201309:
                quarter_days[10] += 1
            elif get_year_month(temp_time_value) <= 201312:
                quarter_days[11] += 1
            elif get_year_month(temp_time_value) <= 201403:
                quarter_days[12] += 1
            elif get_year_month(temp_time_value) <= 201406:
                quarter_days[13] += 1
            elif get_year_month(temp_time_value) <= 201409:
                quarter_days[14] += 1
            elif get_year_month(temp_time_value) <= 201412:
                quarter_days[15] += 1
        sum_index = sum_index + temp_quarter_16
        time_16_days.append(quarter_days)
    else:
        temp_quarter_16 = len(stocks)-stock_total_indexs[stock_index]
        for quarter_index in xrange(temp_quarter_16):
            temp_time_value = times[sum_index+quarter_index]
            if get_year_month(temp_time_value) <= 201103:
                quarter_days[0] += 1
            elif get_year_month(temp_time_value) <= 201106:
                quarter_days[1] += 1
            elif get_year_month(temp_time_value) <= 201109:
                quarter_days[2] += 1
            elif get_year_month(temp_time_value) <= 201112:
                quarter_days[3] += 1
            elif get_year_month(temp_time_value) <= 201203:
                quarter_days[4] += 1
            elif get_year_month(temp_time_value) <= 201206:
                quarter_days[5] += 1
            elif get_year_month(temp_time_value) <= 201209:
                quarter_days[6] += 1
            elif get_year_month(temp_time_value) <= 201212:
                quarter_days[7] += 1
            elif get_year_month(temp_time_value) <= 201303:
                quarter_days[8] += 1
            elif get_year_month(temp_time_value) <= 201306:
                quarter_days[9] += 1
            elif get_year_month(temp_time_value) <= 201309:
                quarter_days[10] += 1
            elif get_year_month(temp_time_value) <= 201312:
                quarter_days[11] += 1
            elif get_year_month(temp_time_value) <= 201403:
                quarter_days[12] += 1
            elif get_year_month(temp_time_value) <= 201406:
                quarter_days[13] += 1
            elif get_year_month(temp_time_value) <= 201409:
                quarter_days[14] += 1
            elif get_year_month(temp_time_value) <= 201412:
                quarter_days[15] += 1
        sum_index = sum_index + temp_quarter_16
        time_16_days.append(quarter_days)

#方法二:没有方法一思路清晰，重写，用方法一取而代之
'''
#提取每只股票每个季度天数
#【定义】季度临时变量
temp_quarter = 1
days = 0

for stock_index in xrange(len(stock_names)):
    quarter_index = 0
    days = 0
    
    if get_month(times[stock_total_indexs[stock_index]]) < 4:
        temp_quarter = 1
    elif get_month(times[stock_total_indexs[stock_index]]) < 7:
        temp_quarter = 2
    elif get_month(times[stock_total_indexs[stock_index]]) < 10:
        temp_quarter = 3
    elif get_month(times[stock_total_indexs[stock_index]]) < 13:
        temp_quarter = 4
    
    if stock_index != len(stock_names) -1:
        for i in xrange(stock_total_indexs[stock_index],stock_total_indexs[stock_index + 1]):
            month = get_month(times[i])
            days = days + 1
            if (month > (3*temp_quarter)) or \
               (days > 2 and get_month(times[i-1]) == 12 and month == 1):
                quarter_days[quarter_index] = days-1
                days = 1
                
                if month < 4:
                    temp_quarter = 1
                elif month < 7:
                    temp_quarter = 2
                elif month < 10:
                    temp_quarter = 3
                elif month < 13:
                    temp_quarter = 4
                
                quarter_index = quarter_index + 1
            if i == (stock_total_indexs[stock_index + 1] - 1):
                quarter_days[quarter_index] = days
                quarter_index = quarter_index + 1
        time_16_days.append(quarter_days[0:quarter_index])
#        print quarter_days[0:quarter_index]
#        print quarter_index
    else:
        for i in xrange(stock_total_indexs[stock_index],len(times)):
            month = get_month(times[i])
            days = days + 1
            if (month > (3*temp_quarter)) or \
               (days > 2 and get_month(times[i-1]) == 12 and month == 1) :
                quarter_days[quarter_index] = days-1
                days = 1
                
                if month < 4:
                    temp_quarter = 1
                elif month < 7:
                    temp_quarter = 2
                elif month < 10:
                    temp_quarter = 3 
                elif month < 13:
                    temp_quarter = 4

                quarter_index = quarter_index + 1
            if i == (len(times) - 1):
                quarter_days[quarter_index] = days
                quarter_index = quarter_index + 1
        time_16_days.append(quarter_days[0:quarter_index])
#        print quarter_index
#        print quarter_days[0:quarter_index]
'''
       
#计算股价收益偏度

#【3】所有股票所有季度股价收益偏度
#【定义】存储股价收益偏度

book = xlwt.Workbook(encoding='utf-8',style_compression=0)
file1_xls = book.add_sheet('skewness',cell_overwrite_ok = True) 
file2_xls = book.add_sheet('market_values',cell_overwrite_ok = True)
file3_xls = book.add_sheet('rit_mean',cell_overwrite_ok = True)

skewness = []
market_values_total = []
rit_mean_total = []
stock_index = 0
sum_quarter_days = 0
sum_quarter_days_before = 0

for stock_index in xrange(len(time_16_days)):
    ske = [0]*16
    market_values_each = [0]*16
    rit_mean_each = [0]*16
    ske_index = 0
    for quarter_each in time_16_days[stock_index]:
        sum_quarter_days = sum_quarter_days_before + quarter_each
        sum_3 = 0
        sum_2 = 0 
        sum_rit = 0
        for rit_each in rit[sum_quarter_days_before:sum_quarter_days]:
            sum_3 = sum_3 + rit_each**3
            sum_2 = sum_2 + rit_each**2
            sum_rit = sum_rit + rit_each
        
        #新需求：如果n=0，市值数据=前一季度值
        if quarter_each==0:
            market_values_each[ske_index] = market_values_each[ske_index-1]
        else:
	        market_values_each[ske_index] = marktv[sum_quarter_days-1]
         
        #新需求：如果n=0,rit_mean_each = 0
        if quarter_each == 0:
            rit_mean_each[ske_index] = 0
        else:
            rit_mean_each[ske_index] = sum_rit/quarter_each
        
        sum_quarter_days_before = sum_quarter_days
        #新需求：n=0,1,2  ske=0
        if quarter_each>=0 and quarter_each<=2:
            ske[ske_index] = 0
        else:
            ske[ske_index] = quarter_each*((quarter_each-1)**(0.5))*sum_3/(quarter_each-2)/((sum_2)**(1.5))
        ske_index = ske_index + 1
    skewness.append(ske)
    market_values_total.append(market_values_each)
    rit_mean_total.append(rit_mean_each)
#    print '------------------------------------------------------------'  

#    print stock_index

print "skewness is:\n" + str(skewness)
print "market value is:\n " + str(market_values_total)
print "rit mean is:\n " + str(rit_mean_total)

#输出结果保存成一个excel
file1_xls.write(0,0,'stock types')
file1_xls.write(0,1,'skewness')
for s_indx in xrange(len(skewness)):
    file1_xls.write(s_indx+1,0,stock_names[s_indx])
    for sk in xrange(len(skewness[s_indx])):
        file1_xls.write(s_indx+1,sk+1,skewness[s_indx][sk])
 

file2_xls.write(0,0,'stock types')
file2_xls.write(0,1,'market_values')
for m_index in xrange(len(market_values_total)):
    file2_xls.write(m_index+1,0,stock_names[m_index])
    for mar in xrange(len(market_values_total[m_index])):
        file2_xls.write(m_index+1,mar+1,market_values_total[m_index][mar])
        
file3_xls.write(0,0,'stock types')
file3_xls.write(0,1,'rit mean')
for k_indx in xrange(len(rit_mean_total)):
    file3_xls.write(k_indx+1,0,stock_names[k_indx])
    for rim in xrange(len(rit_mean_total[k_indx])):
        file3_xls.write(k_indx+1,rim+1,rit_mean_total[k_indx][rim])

book.save('./results/result_V5.xls')











