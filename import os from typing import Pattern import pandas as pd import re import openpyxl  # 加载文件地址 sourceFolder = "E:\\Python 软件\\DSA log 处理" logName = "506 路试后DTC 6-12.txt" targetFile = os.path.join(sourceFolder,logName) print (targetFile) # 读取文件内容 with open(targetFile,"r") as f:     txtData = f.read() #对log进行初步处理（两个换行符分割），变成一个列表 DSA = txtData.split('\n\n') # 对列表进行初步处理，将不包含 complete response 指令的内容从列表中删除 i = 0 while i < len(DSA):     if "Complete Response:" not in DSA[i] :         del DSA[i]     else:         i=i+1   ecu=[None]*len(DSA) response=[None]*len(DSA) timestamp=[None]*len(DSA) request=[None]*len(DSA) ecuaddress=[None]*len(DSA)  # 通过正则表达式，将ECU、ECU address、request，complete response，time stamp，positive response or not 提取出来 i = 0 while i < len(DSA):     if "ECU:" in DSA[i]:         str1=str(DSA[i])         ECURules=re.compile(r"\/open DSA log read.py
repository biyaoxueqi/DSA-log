import os
from typing import Pattern
import pandas as pd
import re
import openpyxl

# 加载文件地址
sourceFolder = "E:\\Python 软件\\DSA log 处理"
logName = "506 路试后DTC 6-12.txt"
targetFile = os.path.join(sourceFolder,logName)
print (targetFile)
# 读取文件内容
with open(targetFile,"r") as f:
    txtData = f.read()
#对log进行初步处理（两个换行符分割），变成一个列表
DSA = txtData.split('\n\n')
# 对列表进行初步处理，将不包含 complete response 指令的内容从列表中删除
i = 0
while i < len(DSA):
    if "Complete Response:" not in DSA[i] :
        del DSA[i]
    else:
        i=i+1


ecu=[None]*len(DSA)
response=[None]*len(DSA)
timestamp=[None]*len(DSA)
request=[None]*len(DSA)
ecuaddress=[None]*len(DSA)

# 通过正则表达式，将ECU、ECU address、request，complete response，time stamp，positive response or not 提取出来
i = 0
while i < len(DSA):
    if "ECU:" in DSA[i]:
        str1=str(DSA[i])
        ECURules=re.compile(r"\/(\w*\s?\w*)\_APP")
        RespRules=re.compile(r'[Complete\sResponse:\s](\w{4}\s[5-7]\w\s\w\w\s\w\w[\s?\w?\w?]*)\n')
        TimeRules=re.compile(r'(\d{9})\n')
        ECUaddressRules=re.compile(r"[Complete\sResponse:\s](1\w\w\d\s)")
        RequestRules = re.compile(r'[Tester\-\>\s](\w{4}\s[1-3]\w\s\w\w\s\w\w\s?\w?\w?\s?\w?\w?)')
        ecu[i]=ECURules.search(str1).group(1)
        response[i]=RespRules.search(str1).group(1)
        timestamp[i]=TimeRules.search(str1).group(1)
        ecuaddress[i]=ECUaddressRules.search(str1).group(1)
        request[i]=RequestRules.search(str1).group(1)
        i=i+1

    else:
        str1 = str(DSA[i])
        ECURules = re.compile(r'\s(\w*\s?\w*)\_APP')
        RespRules = re.compile(r'[Complete\sResponse:\s](\w{4}\s[5-9]\w\s\w\w\s\w\w[\s?\w?\w?]*)\n')
        TimeRules = re.compile(r'(\d{9})\n')
        ECUaddressRules = re.compile(r"[Complete\sResponse:\s](1\w\w\d\s)")
        RequestRules = re.compile(r'[Tester\-\>\s](\w{4}\s[1-3]\w\s\w\w\s\w\w[\s?\w?\w?]*)')
        ecu[i] = ECURules.search(str1).group(1)
        response[i] = RespRules.search(str1).group(1)
        timestamp[i] = TimeRules.search(str1).group(1)
        ecuaddress[i] = ECUaddressRules.search(str1).group(1)
        request[i]=RequestRules.search(str1).group(1)

        i=i+1

DSAData = pd.DataFrame.from_dict(dict([("ECU",ecu),("ECU address",ecuaddress),("request",request),("response",response)]))


print(DSAData)
DSAData.to_excel('DSAlog.xlsx')

# flag=[0]*8
# oder = [None]*len(DSAData)
# expectresult = [None]*len(DSAData)
# result = [None]*len(DSAData)
# status = [None]*len(DSAData)
# i = 0
# while i < len(DSAData):
#     if '1A01 2F DD 0A 03 01' in DSAData['request'][i]:
#         oder[i] = 'set usgmode to inactive'
#
#         if '1A01 6F DD 0A 03 01' in DSAData['response'][i]:
#             status[i] = 'inactive'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[0] = i
#         i=i+1
#     elif '1A01 2F DD 0A 03 02' in DSAData['request'][i]:
#         oder[i] = 'set usgmode to convenience'
#         if '1A01 6F DD 0A 03 02' in DSAData['response'][i]:
#             status[i] = 'convenience'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[1] = i
#         i = i + 1
#     elif '1A01 2F DD 0A 03 0B' in DSAData['request'][i]:
#         oder[i] = 'set usgmode to active'
#         if '1A01 6F DD 0A 03 0B' in DSAData['response'][i]:
#             status[i] = 'active'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[3]= i
#         i = i + 1
#     elif '1A01 2F DD 0A 03 0D' in DSAData['request'][i]:
#         oder[i] = 'set usgmode to driving'
#         if '1A01 6F DD 0A 03 0D' in DSAData['response'][i]:
#             status[i] = 'driving'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[2]= i
#         i = i + 1
#     elif '1A01 2F D1 34 03 00' in DSAData['request'][i]:
#         oder[i] = 'set carmode to normal'
#         if '1A01 6F D1 34 03 00' in DSAData['response'][i]:
#             status[i] = 'normal'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[7] = i
#         i = i + 1
#     elif '1A01 2F D1 34 03 01' in DSAData['request'][i]:
#         oder[i] = 'set carmode to transport'
#         if '1A01 6F D1 34 03 01' in DSAData['response'][i]:
#             status[i] = 'transport'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[4]= i
#         i = i + 1
#     elif '1A01 2F D1 34 03 02' in DSAData['request'][i]:
#         oder[i] = 'set carmode to factory'
#         if '1A01 6F D1 34 03 02' in DSAData['response'][i]:
#             status[i] = 'factory'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[5]= i
#         i = i + 1
#     elif '1A01 2F D1 34 03 05' in DSAData['request'][i]:
#         oder[i] = 'set carmode to Dyno'
#         if '1A01 6F D1 34 03 05' in DSAData['response'][i]:
#             status[i] = 'Dyno'
#             result[i] = 'OK'
#         else:
#             status[i] = 'unknown'
#             result[i] = 'NOK'
#         flag[6]= i
#         i = i + 1
#     elif '1FFF 22 DD 0A' in DSAData['request'][i]:
#         oder[i] = 'read usgmode'
#         if '62 DD 0A 01' in DSAData['response'][i]:
#             status[i] = 'inactive'
#         elif '62 DD 0A 02' in DSAData['response'][i]:
#             status[i] = 'convenience'
#         elif '62 DD 0A 0B' in DSAData['response'][i]:
#             status[i] = 'active'
#         elif '62 DD 0A 0D' in DSAData['response'][i]:
#             status[i] = 'driving'
#         elif '7F 22' in DSAData['response'][i]:
#             status[i] = 'nagtive response'
#             result[i] = 'NOK'
#         else:
#             status [i] = 'unknown'
#             result[i] = 'NOK'
#         i = i + 1
#     elif '1FFF 22 D1 34' in DSAData['request'][i]:
#         oder[i] = 'read carmode'
#         if '62 D1 34 00' in DSAData['response'][i]:
#             status[i] = 'normal'
#         elif '62 D1 34 01' in DSAData['response'][i]:
#             status[i] = 'transport'
#         elif '62 D1 34 02' in DSAData['response'][i]:
#             status[i] = 'factory'
#         elif '62 D1 34 05' in DSAData['response'][i]:
#             status[i] = 'Dyno'
#         elif '7F 22' in DSAData['response'][i]:
#             status[i] = 'nagtive response'
#             result[i] = 'NOK'
#         else:
#             status [i] = 'unknown'
#             result[i] = 'NOK'
#         i = i + 1
#     else:
#         oder[i] = 'unknown'
#         i = i + 1
#
# i = 0
# while i < len(result):
#     # if  result[i] is None:
#         if i < flag[0]:
#             expectresult[i] = 'norespect'
#             result[i] = 'no result'
#         elif i < flag[1]:
#             expectresult[i] = "inactive"
#             if 'inactive' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i] = 'NOK'
#         elif i < flag[2]:
#             expectresult[i] = 'convenience'
#             if 'convenience' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i]= 'NOK'
#             # i = i+1
#         elif i < flag[3]:
#             expectresult[i] = 'driving'
#             if 'driving' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i]= 'NOK'
#             # i = i + 1
#         elif i < flag[4]:
#             expectresult[i] = 'active'
#             if 'active' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i]= 'NOK'
#             # i = i+1
#         elif i < flag[5]:
#             expectresult[i] = 'transport'
#             if 'transport' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i]= 'NOK'
#             # i = i+1
#         elif i < flag[6]:
#             expectresult[i] = 'factory'
#             if 'factory' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i]= 'NOK'
#             # i = i+1
#         elif i < flag[7]:
#             expectresult[i] = 'Dyno'
#             if 'Dyno' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i]= 'NOK'
#         else:
#             expectresult[i] = 'normal'
#             if 'normal' in status[i]:
#                 result[i] = 'OK'
#             else:
#                 result[i] = 'NOK'
#
#             # i = i+1
#         i = i+1
#     # else:
#     #     i=i+1
#
# finalResult = pd.DataFrame.from_dict(dict([('ECU',DSAData['ECU']),('ECU address',DSAData['ECU address']),("REQ",DSAData['request']),('request',oder),("RESP",DSAData['response']),('expect result',expectresult),('response',status),('result',result)]))
# finalResult.to_excel("finalresult.xlsx")
# # failresult = finalResult.sub(other= "DataFrame",axis= 0,fill_value="NOK",)
# # x=pd.DataFrame.
#
# print(flag)
# print(oder)
# print(status)
# print(result)
# print(finalResult)
# # print(x)











