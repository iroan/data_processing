'''
需求：把carbrand.xlsx中的数据集统计成这样的数据
carbrand = {
    'carbrand':{'series':list,'abbr:str'}
    ...
}

实现：
1. 遍历所有数据
    1. 如果车品牌已经在字典中，把B列数据添加到series中
    1. 如果车品牌不在字典中，把A中的字段作为key，把B的值追加到series中，C中的值赋值给abbr

'''

from openpyxl import load_workbook
from pprint import pprint
import json
filepath = 'carbrand.xlsx'
wb = load_workbook(filepath)
sheet = wb['Sheet1']
result_dict = dict()

for index, item in enumerate(zip(sheet['A'],sheet['B'],sheet['C'])):
    brand = item[0].value
    series = item[1].value
    abbr = item[2].value
    if brand in result_dict:
        result_dict[brand]['series'].append(series)
    else:
        tmp = {
            'series':[series],
            'abbr':abbr
        }
        result_dict[brand] = tmp

pprint(result_dict)