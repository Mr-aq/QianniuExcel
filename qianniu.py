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
import os
import random
from time import sleep

from selenium import webdriver
from selenium.common import ElementClickInterceptedException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait

from excel import Excel
from log import Log


class Qianniu:
    excel = Excel()
    log = Log()

    # 下载路径
    download = ''

    def __init__(self):
        self.driver = None

    def create_driver(self, time):
        """
        创建driver
        :return: driver
        """
        chrome_options = webdriver.ChromeOptions()
        prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': self.download}
        self.log.info(messages=f'下载位置{self.download}')
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_argument('--ignore-certificate-errors')  # 忽略CERT证书错误
        chrome_options.add_argument('--ignore-ssl-errors')  # 忽略SSL错误

        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.set_window_position(random.randrange(10, 1000, 100), random.randrange(10, 300, 50))
        # 通过浏览器的dev_tool在get页面钱将.webdriver属性改为"undefined"
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
            "source": """Object.defineProperty(navigator, 'webdriver', {get: () => undefined})""",
        })

        url = 'https://loginmyseller.taobao.com/?from=&f=top&style=&sub=true&redirect_url=https%3A%2F%2Fmyseller.taobao.com%2Fhome.htm%2FQnworkbenchHome%2F'
        sleep(time)
        self.driver.get(url)

    def login(self, account, passwd):
        """
        登录
        :param account: 账号
        :param passwd: 密码
        :return:
        """

        # 切换窗口
        self.driver.switch_to.frame('alibaba-login-box')

        # 输入账号密码
        self.driver.find_element(By.ID, 'fm-login-id').send_keys(account)
        self.driver.find_element(By.ID, 'fm-login-password').send_keys(passwd)
        # 点击登录
        self.driver.find_element(By.XPATH, '//div[@class="fm-btn"]/button').click()

        # 验证账号是否存在
        account_state = '正常'
        sleep(3)
        try:
            error = self.driver.find_element(By.XPATH, '//*[@id="login-error"]/div').text
            if error == '账号名或登录密码不正确':
                account_state = error
                self.log.info(messages=f'{account}--账号名或登录密码不正确')
        except Exception:
            pass

        try:
            error = self.driver.find_element(By.XPATH, '//*[@id="login-error"]/div').text
            if error == '该子账号已经离职':
                account_state = error
                self.log.info(messages=f'{account}--该子账号已经离职')
        except Exception:
            pass

        try:
            self.driver.find_element(By.XPATH, '//*[@id="ice-container"]/div/div/div/div/div[1]/button')
            # WebDriverWait(self.driver, 10, 0.5).until(EC.visibility_of_element_located(
            #     (By.XPATH, '//*[@id="ice-container"]/div/div/div/div/div[1]/button')))
        except Exception:
            pass
        else:
            account_state = '账号需要验证'
            self.log.info(messages=f'{account}--账号需要验证')

        try:
            # 滑动验证码
            sleep(2)
            page = self.driver.page_source
            slider = self.driver.find_element(By.XPATH, '/html/body/center/div[1]/div/div/div/div[2]/span').text
            self.log.info(messages=f'{account}--出现滑动验证码')
            if slider == '向右滑动验证':
                self.crack(slider, 390.4)
                # 点击登录
                self.driver.find_element(By.XPATH, '//div[@class="fm-btn"]/button').click()
        except Exception:
            pass

        # 账号正常进行下一步，账号错误直接关闭
        if account_state == '正常':
            self.log.info(messages=f'{account}--账号正常')
            # 验证码
            try:
                # 判断是否有验证码
                WebDriverWait(self.driver, 30, 0.5).until(
                    EC.visibility_of_element_located((By.XPATH, '//*[@id="btn-submit"]')))
                # sleep(2)
                # self.driver.find_element(By.XPATH, '//*[@id="J_GetCode"]').click()
                self.log.info(messages=f'{account}--输入验证码')
                sleep(5)

                # # 是否点击确定进入主页
                # try:
                #     name = self.driver.find_element(By.XPATH, '//*[@id="icestarkNode"]/div/div/div[2]/div[2]/div/div[1]/div/div[2]/div[1]/div[1]/div[1]').text
                #     # WebDriverWait(self.driver, 60, 0.5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="icestark-container"]/div[1]/div/div[6]/div/div[1]/div[1]')))
                # except Exception:
                #     self.log.info(messages='验证失败')
                #     sleep(60)
            except Exception:
                self.log.info(messages=f'{account}--无需验证码')

            # # 是否进入主页
            # while True:
            #     title = self.driver.title
            #     if title == '千牛商家工作台':
            #         print(title)
            #         self.log.info(messages=f'{account}--登录成功')
            #         break
            try:
                WebDriverWait(self.driver, 1200, 0.5).until(EC.visibility_of_element_located(
                    (By.XPATH, '/html/body/div[1]/div[3]/div[3]/div/div[2]/div/div/a[1]/span')))
            except Exception:
                self.log.info(messages=f'{account}--登录失败')
                account_state = '登录失败'
            else:
                self.log.info(messages=f'{account}--登录成功')
                account_state = '正常'
        return account_state

    def get_info(self, account, start_time, end_time, account_state):
        """
        解析生意参谋数据
        :param account: 账号
        :param start_time: 起始时间
        :param end_time: 结束时间
        :param account_state: 账号状态
        :return:
        """
        if account_state != '正常':
            data_state = '数据权限未开通'
            order_state = '订单权限未开通'
            return 0, data_state, order_state, account_state

        # 等待元素加载
        page = self.driver.page_source
        # sleep(5)
        try:
            WebDriverWait(self.driver, 300, 0.5).until(EC.visibility_of_element_located((By.XPATH,
                                                                                         '/html/body/div[1]/div[3]/div[2]/div/div/div/div/div/div/div[2]/div[2]/div/div[1]/div/div[2]/div[1]/div[1]/div[1]')))
        except Exception:
            self.log.info(messages=f'{account}--主页加载失败')
            data_state = '数据加载失败'
            order_state = '订单加载失败'
        else:
            self.log.info(messages=f'{account}--主页加载成功')

        # 数据权限
        try:
            # 两天金额
            yes_money = self.driver.find_elements(By.XPATH,
                                                  '//*[@id="icestarkNode"]/div/div/div[2]/div[1]/div/div[3]/div/div[2]/div/div[1]/a/div/div/div[3]/div/span')[
                1].text
            tod_money = self.driver.find_elements(By.XPATH,
                                                  '//*[@id="icestarkNode"]/div/div/div[2]/div[1]/div/div[3]/div/div[2]/div/div[1]/a/div/div/div[2]/span')
            a = tod_money[0].text
            b = tod_money[1].text
            tod_money = a + b

            # 处理千位
            if ',' in yes_money:
                yes_money = float(''.join(yes_money.split(',')))
            if ',' in tod_money:
                tod_money = float(''.join(tod_money.split(',')))

            data_state = '正常'
            self.log.info(messages=f'{account}--数据获取成功')
        except Exception:
            self.log.info(messages=f'{account}--数据权限未开通')
            data_state = '数据权限未开通'
            yes_money = 0
            tod_money = 0

        # 总交易额
        total_money = float(yes_money) + float(tod_money)
        self.log.info(messages=f'{account}--总交易额：{total_money}')

        # 交易
        sleep(2)
        # page = self.driver.page_source
        try:
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[3]/div/div[2]/div/div/a[3]/span').click()
        except Exception:
            self.log.info(messages='点击交易按钮失败，正在重试')
            sleep(2)
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div[3]/div/div[2]/div/div/a[3]/span').click()
        sleep(5)
        page = self.driver.page_source
        # 交易权限
        try:
            mess = self.driver.find_element(By.CLASS_NAME, 'next-message-content').text
        except Exception:
            self.log.info(messages=f'{account}--交易页面加载成功')
            self.log.info(messages=f'{account}--订单获取成功')
            order_state = '正常'
        else:
            self.log.info(messages=f'{account}--订单权限未开通')
            order_state = '订单权限未开通'

        # 展开筛选
        WebDriverWait(self.driver, 300, 0.5).until(EC.visibility_of_element_located(
            (By.XPATH, '//*[@id="icestarkNode"]/div/div[3]/div[2]/div/form/div/div[2]/div/div[2]')))
        try:
            self.driver.find_elements(By.CLASS_NAME, 'search-form_foldable-cursor__zLgZq')[1].click()
        except ElementClickInterceptedException:
            account_state = '账号需要二次验证'
            data_state = '二次验证失败'
            order_state = '二次验证失败'
            return 0, data_state, order_state, account_state
        sleep(1)
        WebDriverWait(self.driver, 300, 0.5).until(EC.visibility_of_element_located((By.ID, 'paymentDate')))
        self.driver.find_element(By.ID, 'paymentDate').click()
        sleep(1)
        # 筛选时间
        self.driver.find_element(By.XPATH,
                                 '//*[@id="qn-worbench-container"]/div[2]/div/div[1]/div/span[1]/input').send_keys(
            start_time)
        self.driver.find_element(By.XPATH, '//div[@class="next-range-picker-panel-input"]/span[4]/input').send_keys(
            end_time)
        # 单击确定
        self.driver.find_element(By.XPATH, '//*[@id="qn-worbench-container"]/div[2]/div/div[3]/button[2]').click()
        sleep(1)

        self.log.info(messages=f'{account}--筛选时间为{start_time}-{end_time}')

        # 搜索订单
        self.driver.find_element(By.XPATH, '//*[@id="icestarkNode"]/div/div[3]/div[2]/div/form/div/div[2]/div/button[1]').click()
        sleep(1)
        # 批量导出
        try:
            self.driver.find_element(By.XPATH,
                                     '//div[@class="next-box search-form_box-button-list__+Yi18"]/button[@class="next-btn next-medium next-btn-normal"]').click()
        except Exception:
            self.driver.find_element(By.XPATH,
                                     '//*[@id="icestarkNode"]/div/div[3]/div[2]/div/form/div/div[2]/div/button[2]/span').click()
        sleep(3)
        # 生成报表
        self.driver.find_element(By.XPATH,
                                 '//*[@id="qn-worbench-container"]/div[2]/div[2]/div/div/div[3]/div/button[2]').click()
        sleep(1)
        self.driver.find_elements(By.XPATH,
                                  '//*[@id="qn-worbench-container"]/div[3]/div[2]/div/div/div[2]/div/button')[1].click()

        # 下载报表
        sleep(70)
        # 切换窗口
        wins = self.driver.window_handles
        self.driver.switch_to.window(wins[-1])

        try:
            self.driver.find_element(By.XPATH,
                                     '//*[@id="icestarkNode"]/div/div[2]/div[1]/div[3]/div/div/button').click()
            sleep(5)
        except Exception:
            account_state = '5分钟内只能生成一次报表'
            data_state = '5分钟内只能生成一次报表'
            order_state = '5分钟内只能生成一次报表'
            self.log.info(messages=f'{account}--5分钟内只能生成一次报表')
        else:
            self.log.info(messages=f'{account}--下载报表成功')

        return total_money, data_state, order_state, account_state

    def parse_excel(self, filename, account, passwd, sum_money, account_state, order_state, data_state):
        """
        解析excel文件
        :param filename: 文件路径
        :param account: 账号
        :param passwd: 密码
        :param sum_money: 总金额
        :param account_state: 账号状态
        :param order_state: 订单状态
        :param data_state: 数据状态
        :return:
        """
        file = ''
        if account_state == '正常':
            # 获取最新下载的文件
            sleep(5)
            file_lists = os.listdir(filename)
            file_lists.sort(
                key=lambda fn: os.path.getmtime(filename + "\\" + fn) if not os.path.isdir(filename + "\\" + fn) else 0)
            # 文件完整路径
            file = os.path.join(filename, file_lists[-1])
            # 未下载完
            if '.crdownload' in file:
                sleep(5)
                file = file[:-11]

            self.log.info(messages=f'{account}--报表路径：{file}')
        # 筛选订单
        self.excel.filter(file, account, passwd, sum_money, account_state, order_state, data_state)

    def multi_pro(self, time, location, account, passwd, start_time, end_time):
        # 创建实例
        self.create_driver(time)

        # 批量登录
        account_state = self.login(account=account, passwd=passwd)

        # 批量解析页面
        sum_money, data_state, order_state, account_state = self.get_info(account=account, start_time=start_time,
                                                                          end_time=end_time,
                                                                          account_state=account_state)
        self.parse_excel(filename=location, account=account, passwd=passwd, sum_money=sum_money,
                         order_state=order_state, data_state=data_state, account_state=account_state)

        # 关闭网页
        self.driver.quit()

    def get_track(self, distance):
        """
        根据偏移量获取移动轨迹
        :param distance:
        :return:
        """
        # 移动轨迹
        track_x = []
        track_y = []
        # 当前位移
        current = 0
        # 减速阈值
        mid = distance * 4 / 5
        # 计算间隔
        t = 0.2
        # 初速度
        v = 0

        a1 = random.randint(5, 8)

        while current < distance:
            if current < mid:
                # 加速度为正2
                a = a1
            else:
                # 加速度为负3
                a = -3
            # 初速度v0
            v0 = v
            # 当前速度v = v0 + at
            v = v0 + a * t
            # 移动距离x = v0t + 1/2 * a * t^2
            move = v0 * t + 1 / 2 * a * t * t
            # 当前位移
            current += move
            # 加入轨迹
            track_x.append(round(move))
            # 随机y
            track_y.append(random.uniform(0, 5))

        return track_x, track_y

    # 按照轨迹拖动滑块
    def move_to_gap(self, slider, track_x, track_y):
        ActionChains(self.driver).click_and_hold(slider).perform()
        for i in range(len(track_x)):
            ActionChains(self.driver).move_by_offset(xoffset=track_x[i], yoffset=track_y[i]).perform()
        sleep(0.5)
        ActionChains(self.driver).release().perform()

    def crack(self, slider, distance):
        # 获取移动轨迹
        track_x, track_y = self.get_track(distance)
        # print(track_x, track_y)
        # 拖动滑块
        self.move_to_gap(slider, track_x, track_y)
