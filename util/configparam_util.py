# _*_ coding: utf-8 _*_

import configparser
import os
"""
从配置文件中获取参数
"""

class ConfigEngine(object):
    @staticmethod
    def get_param(file_path,section,key):
        print("%s>%s>%s"%(file_path,section,key))
        config = configparser.ConfigParser()
        config.read(file_path)
        return config.get(section,key)

    @staticmethod
    def get_param_default(section,key):
        file_path = os.path.dirname(os.path.abspath("."))+"/InterfaceAutomation/conf/config.ini"
        config = configparser.ConfigParser()
        config.read(file_path)
        return config.get(section, key)