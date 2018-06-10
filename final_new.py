from openpyxl import load_workbook
from pprint import pprint
import json
filepath = '/home/iroan/data/final_new.xlsx'

def task_0609(filename):
    '''
    功能：提取出每一行数据中M列对应的汽车品牌在N列中出现的汽车集合
    :param filename: 处理的文件名
    :return: 一个json格式的数据
        数据形式：
                [
                    {"品牌":[汽车X,...]}
                    ...
                ]
    '''
    wb = load_workbook(filepath)
    sheet = wb['Sheet1']
    car_brand_set = set()
    result_dict = dict()

    for index,item in enumerate(sheet['M']):
        if index == 1:
            continue
        if item.value != None:
            car_brand_set.add(item.value)
    for i in car_brand_set:
        result_dict[i] = set()

    for index,item in enumerate(zip(sheet['M'],sheet['N'])):
        A = item[0].value
        C = item[1].value
        if A in result_dict:
            result_dict[A].add(C)
    for i in result_dict:
        result_dict[i] = list(result_dict[i])


    pprint(result_dict)

if __name__ == '__main__':
    task_0609(filepath)
