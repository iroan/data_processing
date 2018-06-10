from openpyxl import load_workbook
from pprint import pprint
import json
filepath = '/home/iroan/data/carnum.xlsx'

def task_0609(filename):
    '''
    功能：提取出每一行数据中A字段对应的省份在C中出现的字母集合
    :param filename: 处理的文件名
    :return: 一个json格式的数据
        数据形式：
                [
                    {"省份X":'字幕X','...'}
                    ...
                ]

    实现步骤:
        1. 获取到省份列，字母列数据
        1. 获取省份的集合province_set
        1. 根据集合province_set的字段创建字典
        3. 判断字母是否在 province_set中，如果在，添加到集合中
    '''
    wb = load_workbook(filepath)
    sheet = wb['Sheet1']
    province_set = set()
    result_dict = dict()

    def get_rule_data():
        for index,item in enumerate(sheet['A']):
            if index == 1:
                continue
            if item.value != None:
                province_set.add(item.value)
        for i in province_set:
            result_dict[i] = set()

        for index,item in enumerate(zip(sheet['A'],sheet['C'])):
            A = item[0].value
            C = item[1].value
            if A in result_dict:
                result_dict[A].add(C)

    def get_norule_data():
        for index, item in enumerate(zip(sheet['G'],sheet['I'],sheet['K'])):
            if index == 0:
                continue
            if index == 1:
                for i in item:
                    result_dict[i.value[:-3]] = set()
                continue

            if index < 24:
                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value
                    if '皖' in tmp:
                        result_dict['安徽省'].add(i.value[-1])
                    if '鲁' in tmp:
                        result_dict['山东省'].add(i.value[-1])
                    if '新' in tmp:
                        result_dict['新疆维吾尔自治区'].add(i.value[-1])

            if index == 24:
                # print('{}{}{}'.format(item[0].value,item[1].value,item[2].value))
                for i in item:
                    result_dict[i.value[:-3]] = set()
                continue


            if 24< index < 39:
                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value
                    if '苏' in tmp:
                        result_dict['江苏省'].add(i.value[-1])
                    if '浙' in tmp:
                        result_dict['浙江省'].add(i.value[-1])
                    if '赣' in tmp:
                        result_dict['江西省'].add(i.value[-1])

            if index == 38:
                for i in item:
                    result_dict[i.value[:-3]] = set()
                continue

            if 38< index < 57:
                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value
                    if '鄂' in tmp:
                        result_dict['湖北省'].add(i.value[-1])
                    if '桂' in tmp:
                        result_dict['广西壮族自治区'].add(i.value[-1])
                    if '甘' in tmp:
                        result_dict['甘肃省'].add(i.value[-1])

            if index == 56:
                for i in item:
                    result_dict[i.value[:-3]] = set()
                continue

            if 56< index < 69:
                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value
                    if '晋' in tmp:
                        result_dict['山西省'].add(i.value[-1])
                    if '蒙' in tmp:
                        result_dict['内蒙古自治区'].add(i.value[-1])
                    if '陕' in tmp:
                        result_dict['陕西省'].add(i.value[-1])

            if index == 69:
                for i in item:
                    result_dict[i.value[:-3]] = set()
                continue

            if 70< index < 80:
                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value

                    if '吉' in tmp:
                        result_dict['吉林省'].add(i.value[-1])
                    if '闽' in tmp:
                        result_dict['福建省'].add(i.value[-1])
                    if '贵' in tmp:
                        result_dict['贵州省'].add(i.value[-1])

            if index == 80:
                for i in item:
                    result_dict[i.value[:-3]] = set()
                continue

            if 80 < index < 105:
                if index == 89:
                    result_dict[item[2].value[:-3]] = set()
                # TODO: 缺一个F,记得要补上
                if index == 99:
                    result_dict[item[2].value[:-3]] = set()

                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value

                    if '藏' in tmp:
                        result_dict['西藏自治区'].add(i.value[-1])
                    if '琼' in tmp:
                        result_dict['海南省'].add(i.value[-1])
                    if '粤' in tmp:
                        result_dict['广东省'].add(i.value[-1])
                    if '川' in tmp:
                        result_dict['四川省'].add(i.value[-1])
                    if '青' in tmp:
                        result_dict['青海省'].add(i.value[-1])

            if index == 105:
                for index, i in enumerate(item):
                    if index == 2:
                        continue
                    result_dict[i.value[:-3]] = set()
                continue

            if 105 < index < 112:
                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value

                    if '宁' in tmp:
                        result_dict['宁夏回族自治区'].add(i.value[-1])
                    if '渝' in tmp:
                        result_dict['重庆市'].add(i.value[-1])

            if index == 112:
                for index, i in enumerate(item):
                    result_dict[i.value[:-3]] = set()
                continue

            if 112 < index:
                for i in item:
                    if i.value == None:
                        continue
                    tmp = i.value

                    if '京' in tmp:
                        result_dict['北京市'].add(i.value[-1])
                    if '津' in tmp:
                        result_dict['天津市'].add(i.value[-1])
                    if '沪' in tmp:
                        result_dict['上海市'].add(i.value[-1])

    get_rule_data()
    get_norule_data()
    for index,i in enumerate(result_dict):
        result_dict[i] = list(result_dict[i])

    pprint(result_dict)

if __name__ == '__main__':
    task_0609(filepath)
