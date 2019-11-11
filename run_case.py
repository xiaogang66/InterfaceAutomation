# _*_ coding: utf-8 _*_


"""
请求基础处理类（数据依赖处理、执行请求、结果解析）
"""
from util.request_util import  RequestUtil
from util.excel_util import ExcelUtil
from util.data_util import DataUtil
from util.assert_util import AssertUtil
from util.log_util import Logger
from requests.cookies import RequestsCookieJar
import re



class RunCase(object):

    CASE_ID = 1
    MODULE_NAME = 2
    CASE_NAME = 3
    RUN_FLAG = 4
    URL = 5
    REQUEST_METHOD = 6
    HEADERS = 7
    COOKIES = 8
    REQUEST_PARAM= 9
    EXP_RESULT = 10
    STATUS_CODE = 11
    RESPONSE_TEXT = 12
    ASSET_TYPE = 13
    ASSET_PATTERN = 14
    EXEC_RESLT = 15

    """定义常量，指定表格每一列"""

    def __init__(self,file_name,sheet_name=None,sheet_index=0):
        self.requestUtil = RequestUtil()
        self.excelUtil = ExcelUtil(file_name,sheet_name,sheet_index)
        self.dataUtil = DataUtil()
        self.assetUtil = AssertUtil()
        self.logger = Logger(self.__class__.__name__).get_logger_with_level()
        self.cookie_dict = {}

    def run_case_by_row_no(self,row_no):
        """根据行号执行单个用例"""
        # 按行取数
        row_data = self.excelUtil.get_data_by_row_no(row_no)

        # 数据准备
        case_id = row_data[self.CASE_ID-1]
        # module_name = row_data[self.MODULE_NAME-1]
        run_flag = row_data[self.RUN_FLAG-1]
        if run_flag == '否':
            # 用例不执行
            return
        elif run_flag == '是':
            url = row_data[self.URL-1]
            request_method = row_data[self.REQUEST_METHOD-1]
            # 请求头处理
            headers = row_data[self.HEADERS-1]
            if headers is None:
                headers = {}
            else:
                headers = self.dataUtil.str_to_json(headers)
            # cookie处理
            cookies = row_data[self.COOKIES-1]
            if cookies:
                # 进行cookie的解析处理，判断是否存在cookie依赖
                depend_cookie = self.cookie_depend(cookies)
                if depend_cookie is not None:
                    if type(depend_cookie) == RequestsCookieJar:
                        cookies = depend_cookie
                    else:
                        cookies = self.dataUtil.str_to_json(depend_cookie)
            request_param = row_data[self.REQUEST_PARAM-1]
            if request_param is not None:
                request_param = self.data_depend(request_param)
            exp_result = row_data[self.EXP_RESULT-1]
            asset_type = row_data[self.ASSET_TYPE-1]
            asset_pattern = row_data[self.ASSET_PATTERN-1]

            # 执行并记录结果
            self.logger.info("请求URL：%s" % url)
            self.logger.info("请求参数：%s" % request_param)
            self.logger.info("请求头：%s" % headers)
            self.logger.info("请求cookie：%s" % cookies)
            response = None
            if request_method == 'get':
                response = self.requestUtil.do_get(url,request_param,headers,cookies)
            elif request_method == 'post':
                # 将字符串转换成json对象
                json_param = self.dataUtil.str_to_json(request_param)
                response = self.requestUtil.do_post(url,json_param,'',headers,cookies)
            response_text = response.text.strip()
            if case_id in self.cookie_dict:
                self.cookie_dict[case_id] = response.cookies
            self.logger.info("请求结果：%s\n"%response_text)
            self.excelUtil.set_data_by_row_col_no(row_no,self.STATUS_CODE,response.status_code)
            self.excelUtil.set_data_by_row_col_no(row_no,self.RESPONSE_TEXT,response_text)

            # 断言判断，记录最终结果
            result = self.asset_handle(exp_result,response_text,asset_type,asset_pattern)
            if result:
                self.excelUtil.set_data_by_row_col_no(row_no, self.EXEC_RESLT, 'pass')
            else:
                self.excelUtil.set_data_by_row_col_no(row_no, self.EXEC_RESLT, 'fail')

    def data_depend(self,request_param):
        """处理数据依赖
            ${test_03.data.orderId}   表示对返回结果的部分属性存在依赖
        """
        request_param_final = None
        # 处理返回结果属性依赖
        match_results = re.findall(r'\$\{.+?\..+?\}', request_param)
        if match_results is None or match_results == []:
            return request_param
        else:
            for var_pattern in match_results:
                # 只考虑匹配到一个的情况
                start_index  = var_pattern.index("{")
                end_index  = var_pattern.rindex("}")
                # 得到${}$中的值
                pattern = var_pattern[start_index+1:end_index]
                spilit_index = pattern.index(".")
                # 得到依赖的case_id和属性字段
                case_id = pattern[:spilit_index]
                proper_pattern = pattern[spilit_index+1:]
                row_no = self.excelUtil.get_row_no_by_cell_value(case_id,self.CASE_ID)
                response = self.excelUtil.get_data_by_row_col_no(row_no,self.RESPONSE_TEXT)
                result = self.dataUtil.json_data_analysis(proper_pattern,response)
                # 参数替换，str(result)进行字符串强转，防止找到的为整数
                request_param_final = request_param.replace(var_pattern,str(result),1)
            return request_param_final

    def cookie_depend(self,request_param):
        """处理数据依赖
			1、${test_01}                表示对返回cookie存在依赖
            2、${test_03.data.orderId}   表示对返回结果的部分属性存在依赖
        """
        cookie_final = None
        # 处理对返回cookie的依赖
        match_results = re.match(r'^\$\{(.[^\.]+)\}$', request_param)
        if match_results:
            # 用例返回cookie依赖
            depend_cookie = self.cookie_dict[match_results.group(1)]
            return depend_cookie
        else:
            # 非用例返回cookie依赖
            cookie_final = self.data_depend(request_param)
            return cookie_final


    def asset_handle(self,exp_result,response_text,asset_type,asset_pattern):
        """根据断言方式进行断言判断"""
        asset_flag = None
        if asset_type == '相等':
            if asset_pattern is None or asset_pattern == '':
                asset_flag = self.assetUtil.equals(exp_result,response_text)
            else:
                exp_value = self.dataUtil.json_data_analysis(asset_pattern,exp_result)
                response_value = self.dataUtil.json_data_analysis(asset_pattern,response_text)
                asset_flag = self.assetUtil.equals(exp_value, response_value)
        elif asset_type == '包含':
            asset_flag = self.assetUtil.contains(response_text,asset_pattern)
        elif asset_type == '正则':
            asset_flag = self.assetUtil.re_matches(response_text,asset_pattern)
        return asset_flag