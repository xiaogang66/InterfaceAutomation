# _*_ coding: utf-8 _*_


"""
接口自动化运行入口
"""
from util.configparam_util import ConfigEngine
from util.log_util import Logger
from util.excel_util import ExcelUtil
from run_case import RunCase
import re
import ddt
import unittest
import os
import time
import HTMLTestRunner

file_name = ConfigEngine.get_param_default("caseFileSetting", "caseFile")
sheet_name = ConfigEngine.get_param_default("caseFileSetting", "sheetName")
logger = Logger("RunMain").get_logger_with_level()
runCase = RunCase(file_name, sheet_name)
excelUtil = ExcelUtil(file_name, sheet_name)
excelUtil.load_excel()
data_list = excelUtil.get_case_list()

@ddt.ddt
class RunMain(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.file_name = file_name
        cls.sheet_name = sheet_name
        cls.logger = logger
        cls.runCase = runCase
        cls.excelUtil = excelUtil

    @ddt.data(*data_list)
    def test_run_case(self,data):
        self.excelUtil.load_excel()
        # 获取所有cookie列数据，判断是否存在用例依赖
        cookies = self.excelUtil.get_data_by_col_no(self.runCase.COOKIES)
        # 如果匹配${test_01}成功，则表示存在依赖
        pattern = '^\$\{(.[^\.]+)\}$'
        cookie_list = []
        match_result = None
        for cookie_depend in cookies:
            if cookie_depend is not None:
                match_result = re.match(pattern,cookie_depend)
            if match_result:
                cookie_list.append(match_result.group(1))
        self.runCase.cookie_dict = dict.fromkeys(cookie_list,'')
        result = self.runCase.run_case_by_data(data)
        self.assertTrue(result)


if __name__ == '__main__':
    # 设置报告名称
    report_path = os.path.dirname(os.path.abspath('.')) + '/InterfaceAutomation/reports/'  # 设置报告路径
    now = time.strftime("%Y-%m-%d-%H_%M_%S", time.localtime(time.time()))  # 获取当前时间
    HtmlFile = report_path + now + "HTMLtemplate.html"
    fp = open(HtmlFile, "wb")  # 二进制写文件流
    # 构建suite
    suite = unittest.TestLoader().loadTestsFromTestCase(RunMain)
    runner = HTMLTestRunner.HTMLTestRunner(stream=fp, title=u"xx测试报告", description=u"用例测试情况")
    runner.run(suite)