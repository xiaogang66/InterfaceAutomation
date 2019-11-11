# _*_ coding: utf-8 _*_


"""
接口自动化运行入口
"""
from util.configparam_util import ConfigEngine
from util.log_util import Logger
from util.excel_util import ExcelUtil
from run_case import RunCase
import re

class RunMain(object):

    def __init__(self):
        self.file_name = ConfigEngine.get_param_default("caseFileSetting","caseFile")
        self.sheet_name = ConfigEngine.get_param_default("caseFileSetting", "sheetName")
        self.logger = Logger(self.__class__.__name__).get_logger_with_level()
        self.runCase = RunCase(self.file_name,self.sheet_name)

    def auto_case_exec(self):
        excelUtil = ExcelUtil(self.file_name,self.sheet_name)
        excelUtil.load_excel()
        max_row = excelUtil.sheet.max_row
		# 获取所有cookie列数据，判断是否存在用例依赖
        cookies = excelUtil.get_data_by_col_no(self.runCase.COOKIES)
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
        row_no = 2
        while row_no <= max_row:
            case_id= excelUtil.get_data_by_row_col_no(row_no,self.runCase.CASE_ID)
            module_name = excelUtil.get_data_by_row_col_no(row_no,self.runCase.MODULE_NAME)
            case_name = excelUtil.get_data_by_row_col_no(row_no,self.runCase.CASE_NAME)
            self.logger.info("执行用例：%s-%s-%s" % (case_id,module_name,case_name))
            self.runCase.run_case_by_row_no(row_no)
            row_no = row_no + 1
        self.logger.info("------>全部用例执行成功")


if __name__ == '__main__':
    rm = RunMain()
    rm.auto_case_exec()