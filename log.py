# -*- coding: utf-8 -*-
"""
@Time ： 2024/2/21 17:00
@Auth ： Mr.AQ
@File ：log.py
@IDE ：PyCharm

 * ━━━━━━神兽出没━━━━━━
 * 　　　┏┓　　 ┏┓
 * 　　┏┛┻━━━━┛┻━━┓
 * 　　┃　　　　    ┃
 * 　　┃　　　━　   ┃
 * 　　┃  ┳┛　┗┳   ┃
 * 　　┃　　　　　　 ┃
 * 　　┃　　　┻　　　┃
 * 　　┃　　　　　　 ┃
 * 　　┗━┓　　　┏━━━┛ Code is far away from bug with the animal protecting
 * 　　　　┃　　　┃     神兽保佑,代码无bug
 * 　　　　┃　　　┃
 * 　　　　┃　　　┗━━━━━┓
 * 　　　　┃　　　　　　　┣┓
 * 　　　　┃　　　　　　　┏┛
 * 　　　　┗┓┓┏━━┓┓┏━━━┛
 * 　　　　 ┃┫┫　 ┃┫┫
 * 　　　　 ┗┻┛　 ┗┻┛
"""
import logging


class Log:
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(level=logging.INFO)
        self.logger.handlers.clear()
        self.handler = logging.FileHandler("log.txt", encoding='utf-8')
        self.handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.handler.setFormatter(formatter)
        self.logger.addHandler(self.handler)

    def info(self, messages):
        self.logger.info(messages)

    def debug(self, messages):
        self.logger.debug(messages)

    def error(self, messages):
        self.logger.error(messages)

    def warning(self, messages):
        self.logger.warning(messages)

