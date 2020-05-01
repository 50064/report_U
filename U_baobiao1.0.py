import pandas as pd
import re
import xlrd

def get_str_Ulevel(i,h01_u):     #拿到'10kVⅠ'
    str_station_Ulevel_1 = re.findall('[\a-zA-Z0-9ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+', h01_u['"descr"'].at[i])      #拿到['/10kVⅡ', 'AB']
    str_station_Ulevel_2 = str_station_Ulevel_1[0].split('/')       #拿到['', '10kVⅢ']
    str_station_Ulevel_3 = str_station_Ulevel_2[-1]     #拿到10kVⅠ
    return str_station_Ulevel_3

def get_str_stationname(i,h01_u):       #拿到'裕隆变'
    str_stationname = re.findall('[\u4e00-\u9fa5]+', h01_u['"descr"'].at[i])    #拿到['裕隆变', '10kVⅠ母线AB线电压']
    return str_stationname[0]

def new_sheet(h01_u):       #拿到合并后的表
    h01_u.insert(1, '"descr"', value='')        #插入描述列后得到新的h01_u，即合并后的表
    for j in h02_u.index:
        for i in h01_u.index:
            if h01_u['"id"'].at[i] == h02_u['"id"'].at[j]:
                h01_u['"descr"'].at[i] = h02_u['"descr"'].at[j]
    return h01_u

excel_h_u = pd.read_excel('./h_u.xlsx','附表5 母线电压波动率统计表')
h01_u = pd.read_excel('./01_u.xls')        #Toad导出原始数据
h02_u = pd.read_excel('./02_u.xls')        #Toad导出，为了找站名

h01_u = new_sheet(h01_u)
# h01_u.to_excel('C:/try.xls', index=0)
# pd.options.display.max_columns = 999
# print(h01_u.head())
# print(excel_h_u.head())
def fill(x1,x2):        #填内容用
    for j in range(x1, x2):
        for i in h01_u.index:
            if get_str_stationname(i, h01_u) in str(excel_h_u['变电站'].at[j]) and get_str_Ulevel(i, h01_u) in str(excel_h_u['母线名称'].at[j]):
                excel_h_u['月最大电压（kV）'].at[j] = h01_u['"m_max"'].at[i]
                excel_h_u['最大值发生时间'].at[j] = h01_u['"m_max_t"'].at[i]
                excel_h_u['月最小电压（kV）'].at[j] = h01_u['"m_min"'].at[i]
                excel_h_u['最小值发生时间'].at[j] = h01_u['"m_min_t"'].at[i]
                excel_h_u['平均电压（kV）'].at[j] = h01_u['"average"'].at[i]
                excel_h_u['运行时间（分钟）'].at[j] = h01_u['"regular_time"'].at[i]
                excel_h_u['越上限时间（分钟）'].at[j] = h01_u['"h_time"'].at[i]
                excel_h_u['越下限时间（分钟）'].at[j] = h01_u['"l_time"'].at[i]
                # print(excel_h_u['最大值发生时间'].at[j].dtype)
                # print(h01_u['"m_max_t"'].at[i].dtype)

excel_h_u['最大值发生时间'] = excel_h_u['最大值发生时间'].astype(str)
excel_h_u['最小值发生时间'] = excel_h_u['最小值发生时间'].astype(str)
# type(h01_u['最小值发生时间'])
# print(type(h01_u['月最大电压（kV）'].at[5]))
# print(excel_h_u.columns)
fill(0, 46)     #(表格-2,表格-1)
fill(47, 190)
fill(191, 247)

excel_h_u.to_excel('./h_u_out.xlsx', index=0)

# h01_u.to_excel('C:/try.xls')
