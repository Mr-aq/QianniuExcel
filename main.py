# -*- coding: utf-8 -*-
"""
@Time ： 2024/2/29 15:47
@Auth ： Mr.AQ
@File ：main.py
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
from PyQt5.QtGui import QIcon

import os
import sys
from mainFrame import Ui_MainWindow
from PyQt5.QtWidgets import QApplication, QMainWindow

import multiprocessing

os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2'

if __name__ == "__main__":
    multiprocessing.freeze_support()
    App = QApplication(sys.argv)    # 创建QApplication对象，作为GUI主程序入口
    aw = Ui_MainWindow()    # 创建主窗体对象，实例化Ui_MainWindow
    w = QMainWindow()      # 实例化QMainWindow类
    w.setWindowIcon(QIcon("icon.ico"))  # 设置图标
    aw.setupUi(w)         # 主窗体对象调用setupUi方法，对QMainWindow对象进行设置
    w.show()               # 显示主窗体
    sys.exit(App.exec_())   # 循环中等待退出程序


