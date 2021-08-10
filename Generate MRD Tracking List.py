# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd

"""
1. 将从DT MRD view中导出来的list的存放地址及文件名放入file_path中；
2. 注意：是两个反斜杠"\\"
"""
file_path = 'C:\\Users\\zhang.jing\\Desktop\\Lambda E4.xlsx'
df = pd.read_excel(file_path)
df=df.fillna(value="")

# 将部门Dept和Leader拿出来建立单独的DataFrame，去重，重置索引。
Dept_and_Leader=pd.DataFrame({'ECU Unit':list(df['ECU Unit']), 
                              'SW Leader':list(df['SW Leader'])})
Dept_and_Leader.drop_duplicates(subset=['ECU Unit', 'SW Leader'], 
                                keep='first',inplace=True)
Dept_and_Leader=Dept_and_Leader.reset_index(drop=True)


# 将MRD Delivery Date拿出来，去重，存为list后，进行排序。
MRD_Delivery_Date = list(set(list(df['MRD Delivery Date'])))
MRD_Delivery_Date_sorted=sorted(MRD_Delivery_Date)

Date_ECU_list={}
for date in MRD_Delivery_Date_sorted:
    ECU_str_list=[]
    for i,j in zip(list(Dept_and_Leader['ECU Unit']),list(Dept_and_Leader['SW Leader'])):    
        str_list=''
        for k in range(df.shape[0]):
            if df.loc[k,'ECU Unit']==i and df.loc[k,'SW Leader']==j and str(df.loc[k,'MRD Delivery Date'])==date and df.loc[k,'SCC Status']=='1. Planned':
                str_list=str_list+str(df.loc[k,'ECU Instance'])+' '+str(df.loc[k,'ECU Configuration'])+'; '
        ECU_str_list.append(str_list)
    Date_ECU_list.setdefault(date,'NA')
    Date_ECU_list[date]=ECU_str_list

# 将字典Date_ECU_list转换成DataFrame前，将整列为空的列删除。
null_list=[]
for m in range(Dept_and_Leader.shape[0]):
    null_list.append('')
for key in list(Date_ECU_list.keys()):
    if Date_ECU_list.get(key)==null_list:
        del Date_ECU_list[key]
        
MRD_Tracking=pd.DataFrame(Date_ECU_list)
MRD_Tracking=pd.concat([Dept_and_Leader,MRD_Tracking],axis=1)

"""
将MRD_Tracking存为另外一个单独的excel表格
1. 把MRD跟踪表需存放的地址替代以下地址并命名；
2. 注意：是两个反斜杠"\\"
MRD_Tracking.to_excel('C:\\Users\\zhang.jing\\Desktop\\MRD_Tracking.xlsx')
"""

"""
将MRD_Tracking写在同一excel表格不同sheet里的代码
1. 把MRD跟踪表需存放的地址替代以下地址，将文件名替换以下名为“Lambda E4”的文件名；
2. 注意：是两个反斜杠"\\"
"""
from openpyxl import load_workbook

book = load_workbook('C:\\Users\\zhang.jing\\Desktop\\Lambda E4.xlsx')
writer=pd.ExcelWriter("C:\\Users\\zhang.jing\\Desktop\\Lambda E4.xlsx",engine='openpyxl')
writer.book = book
MRD_Tracking.to_excel(writer,'MRD_Tracking')
writer.save()

