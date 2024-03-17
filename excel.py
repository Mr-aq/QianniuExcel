# -*- coding: utf-8 -*-
"""
@Time ： 2024/2/28 8:50
@Auth ： Mr.AQ
@File ：excel.py
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
import os.path

import pandas as pd

from log import Log


class Excel:
    log = Log()

    def filter(self, file_path, account, passwd, sum_money, account_state, order_state, data_state):
        '''
        对excel进行筛选，选出符合条件的,计算实际金额，并将店铺名和金额存入excel文件
        :return:
        '''
        if account_state == '正常':
            # 读取excel文件
            pd.set_option('display.encoding', 'gbk')
            df_excel = pd.read_excel(file_path, engine='openpyxl')
            # 筛选
            df_filter = df_excel[df_excel['淘特订单'] == '是']
            # 计算实际金额总和
            sum = df_filter['买家实际支付金额'].sum()
            num = len(df_filter['买家实际支付金额'])
            # 存入excel文件
            data = [account, passwd, num, sum, sum_money, account_state, order_state, data_state]
        else:
            data = [account, passwd, 0, 0, 0, account_state, order_state, data_state]

        self.save_excel(data=data)

    def save_excel(self, data):
        # 创建表
        if not os.path.exists('订单计算.xlsx'):
            datas = {
                '子账号': [],
                '登录密码': [],
                '金额': [],
                '订单数量': [],
                '生意参谋金额': [],
                '金额订单(状态)': [],
                '生意参谋(状态)': [],
                '账号(状态)': []
            }

            df = pd.DataFrame(datas)
            df.to_excel(excel_writer='订单计算.xlsx', index=False)
            # 创建一个 ExcelWriter 对象
            # with pd.ExcelWriter('订单计算.xlsx', engine='xlsxwriter') as w:
            #     # 将数据框写入 Excel
            #     df.to_excel(w, index=False, sheet_name='data')
            #     # 获取 ExcelWriter 对象的 worksheet
            #     ws = w.sheets['data']
            #     # 设置列宽
            #     ws.set_column('A:B', 30)
            #     ws.set_column('C:H', 20)

        # 把数据存进表
        df_excel = pd.read_excel('订单计算.xlsx', engine='openpyxl')
        row_count = df_excel.shape[0]

        df_excel.loc[row_count] = data
        df_excel.to_excel(excel_writer='订单计算.xlsx', index=False)
        with pd.ExcelWriter('订单计算.xlsx', engine='xlsxwriter') as w:
            # 将数据框写入 Excel
            df_excel.to_excel(w, index=False, sheet_name='data')
            # 获取 ExcelWriter 对象的 worksheet
            ws = w.sheets['data']
            # 设置列宽
            ws.set_column('A:B', 30)
            ws.set_column('C:H', 20)

        self.log.info(messages=f'{data[0]}--数据存入成功')

    def get_acc(self, file_name):
        '''
        解析excel中的账号
        :return:账号，密码
        '''
        df_excel = pd.read_excel(file_name, engine='openpyxl')
        account = df_excel['子账号'].values.tolist()
        passwd = df_excel['登录密码'].values.tolist()

        self.log.info(messages='账号密码获取成功')

        return account, passwd
