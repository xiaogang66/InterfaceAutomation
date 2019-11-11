# _*_ coding: utf-8 _*_

"""
数据处理类（数据格式转换、json解析）
"""
from jsonpath_rw import jsonpath,parse
import json

class DataUtil(object):

    def json_data_analysis(self,pattern,str_data):
        dict_data = json.loads(str_data)
        json_exe = parse(pattern)
        madle = json_exe.find(dict_data)
        result = [math.value for math in madle]
        if result is None or result == []:
            return None
        else :
            return result[0]

    def str_to_json(self,strs):
        """数据解析成格式化的json格式"""
        result = json.dumps(strs,ensure_ascii=False,sort_keyss=True,indent=2)
        return result

    def json_to_str(self,jsons):
        """数据解析成格式化的json格式"""
        result = json.loads(jsons)
        return result