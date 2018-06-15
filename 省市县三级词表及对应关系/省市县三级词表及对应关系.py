'''
需求： 根据全国省市县.xlsx获取：
    1. 省市县与其对应的简称的映射（字典的key,value关系)
    1. 省到市的映射（字典的key,value关系)
    1. 省到县的映射（字典的key,value关系)
    1. 市到县的映射（字典的key,value关系)
    1. 省词表
    1. 市词表
    1. 县词表

注意：
    1. 台湾省，天津市，澳门特别行政区，重庆市，香港特别行政区
'''

from openpyxl import load_workbook
from pprint import pprint
import json

filepath = '全国省市县.xlsx'
wb = load_workbook(filepath)
sheet = wb['Sheet1']
result_dict = dict()


def abbr():
    abbr_dict = {}
    abbr_province2city = {}
    abbr_city2country = {}
    abbr_province2country = {}
    provinces = list()
    citys = list()
    countrys = list()


    for index, item in enumerate(zip(sheet['A'], sheet['B'], sheet['D'], sheet['E'], sheet['G'], sheet['H'])):
        if index == 0:
            continue
        province_full = item[0].value
        province_abbr = item[1].value
        city_full = item[2].value
        city_abbr = item[3].value
        country_full = item[4].value
        country_abbr = item[5].value

        if province_full in ['台湾省','澳门特别行政区','香港特别行政区']:
            continue
        '''省市县与其对应的简称的映射（字典的key,value关系)'''
        if province_full not in abbr_dict:
            abbr_dict[province_full] = province_abbr
        if city_full not in abbr_dict:
            abbr_dict[city_full] = city_abbr
        if country_full not in abbr_dict:
            abbr_dict[country_full] = country_abbr

        '''省到市的映射（字典的key,value关系)'''
        '''要处理北京市这种没有上级（省）的城市'''
        if province_full[-1] == '市':
            if province_full not in abbr_city2country:
                abbr_city2country[province_full] = {country_full}
            else:
                abbr_city2country[province_full].add(country_full)
        else:
            if province_full not in abbr_province2city:
                abbr_province2city[province_full] = {city_full}
            else:
                abbr_province2city[province_full].add(city_full)

        '''市到县的映射（字典的key,value关系)'''
        if city_full in ['市辖区','县']:
            pass
        else:
            if city_full not in abbr_city2country:
                abbr_city2country[city_full] = {country_full}
            else:
                abbr_city2country[city_full].add(country_full)

        '''省到县的映射（字典的key,value关系)'''
        if province_full[:-1] == '市':
            pass
        else:
            if province_full not in abbr_province2country:
                abbr_province2country[province_full] = {country_full}
            else:
                abbr_province2country[province_full].add(country_full)

    for i in abbr_province2city:
        abbr_province2city[i] = list(abbr_province2city[i])

    for i in abbr_city2country:
        abbr_city2country[i] = list(abbr_city2country[i])

    for i in abbr_province2country:
        abbr_province2country[i] = list(abbr_province2country[i])

    '''获取省词表'''
    provinces = list(abbr_province2city.keys())

    '''获取市词表'''
    citys = list(abbr_city2country.keys())

    '''获取县词表'''
    tmp = []
    for i in abbr_city2country:
        tmp.append(abbr_city2country[i])

    for i in expand_list(tmp):
        countrys.append(i)

    return abbr_dict,abbr_province2city,abbr_city2country,abbr_province2country,provinces,citys,countrys

def expand_list(nests):
    '''传入一个list的嵌套数据，返回一个平坦化的list的生成器'''
    for item in nests:
        if isinstance(item,(list,tuple)):
            for sub_item in expand_list(item):
                yield sub_item
        else:
            yield item


abbr_dict,abbr_province2city,abbr_city2country,abbr_province2country,provinces,citys,countrys = abbr()
pprint(countrys)
