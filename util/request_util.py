"""
请求基础处理类
"""
import requests

class RequestUtil(object):


    def do_get(self,url,params='',headers={},cookies={}):
        """get请求处理"""
        response = requests.get(url,params,verify=False,headers=headers,cookies=cookies)
        return response

    def do_post(self,url,json={},data='',headers={},cookies={}):
        """post请求处理，传入json参数或普通参数"""
        response = None
        if json is not None:
            response = requests.get(url,json=json,verify=False,headers=headers,cookies=cookies)
        elif data is not None:
            response = requests.get(url,data=data,verify=False,heads=headers,cookies=cookies)
        return response