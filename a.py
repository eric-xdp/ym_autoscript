#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @File  : a.py
# @Author: shijiu.Xu
# @Date  : 2020/10/27 
# @SoftWare  : PyCharm

import psutil
import pywinauto
from pywinauto.application import Application
import time, datetime
from pyautogui import hotkey
import pandas as pd
import openpyxl
import logging

LOG_FORMAT = "%(asctime)s %(name)s %(levelname)s ：%(message)s "     # 配置输出日志格式
DATE_FORMAT = '%Y-%m-%d  %H:%M:%S %a '  # 配置输出时间的格式，注意月份和天数不要搞乱了
logging.basicConfig(level=logging.DEBUG, format=LOG_FORMAT, datefmt=DATE_FORMAT,
                    filename=r"send.log"    # 有了filename参数就不会直接输出显示到控制台，而是直接写入文件
                    )


class WeChatRun():

    def __init__(self):
        # self.pid = self.get_pid()
        self.app = Application(backend='uia').connect(process=self.get_pid())
        self.win = self.app[u'微信']

    def get_pid(self):
        for proc in psutil.process_iter():
            try:
                pinfo = proc.as_dict(attrs=['pid', 'name'])
            except psutil.NoSuchProcess:
                pass
            else:
                if 'WeChat.exe' == pinfo['name']: return pinfo['pid']

    def get_search(self, group_name):

        self.search = self.win.child_window(title=u"搜索", control_type="Edit")
        self.search.draw_outline()
        cords = self.search.rectangle()
        pywinauto.mouse.click(button='left', coords=(cords.left + 10, cords.top + 10))
        time.sleep(0.1)
        pywinauto.mouse.click(button='left', coords=(cords.left + 10, cords.top + 10))
        time.sleep(0.2)
        self.win.type_keys(group_name)
        time.sleep(0.5)
        self.win.type_keys('{ENTER}')
        self.win.draw_outline()

    def send_message(self, msg):
        for text in msg.split('\n'):  # line:65
            # 这里用到了另外的一个库，因为用pywinauto 自带的输入模块，表情，空格等是自动略过或者识别不出，达不到按原有缩进样式缩进的效果
            if text.isalnum():
                hotkey('ctrl', 'v')
            else:
                self.win.type_keys(text)
            time.sleep(0.1)
            hotkey('ctrl', 'enter')  # line:67
        hotkey('enter')  # line:68


class ReadExcel():
    """
        封装Pandas对Excel表的处理。自用功能
    """

    def __init__(self, file_path):
        self.path = file_path
        self.book = openpyxl.load_workbook(self.path)
        self.excel_writer = pd.ExcelWriter(self.path, engine='openpyxl')
        self.excel_writer.book = self.book
        # self.excel_writer.sheets = dict((ws.title, ws) for ws in self.book.worksheets)

    # 获取指定的sheet表数据
    def get_data(self, sheet): return pd.read_excel(self.path, sheet_name=sheet).values

    # 添加一个空的sheet
    def add_sheet(self, sheet_name, columns):
        data = pd.DataFrame(columns=columns)
        data.to_excel(self.excel_writer, sheet_name=sheet_name, index=False)
        self.excel_writer.save()

    # 在指定的sheet内追加一行数据。
    def append_data_to_sheet(self, sheet_name, row_list):
        sheet = pd.read_excel(self.excel_writer, sheet_name=sheet_name)
        sheet_data = pd.DataFrame(sheet)
        print(sheet_data)
        sheet_data.loc[sheet_data.shape[0]] = row_list  # 与原数据同格式
        sheet_data.to_excel(self.excel_writer, sheet_name=sheet_name, index=False, header=True)
        self.excel_writer.save()

    # 获取行：获取当天需要通知的行
    def get_row(self, data, weekday, col_name): return data[data[col_name].isin([weekday])]

    # save
    def save(self):
        self.excel_writer.save()
        self.book.save(self.path)

    # close
    def close(self):
        self.excel_writer.close()
        self.book.close()

    # delete
    def delete_sheet(self, sheet_name):
        sheet = self.excel_writer.book[sheet_name]
        self.excel_writer.book.remove(sheet)
        self.save()
class MyRules():

    def __init__(self):
        self.counter = 0
        self.send_hour = [9, 13, 16, 18, 21]
        self.now_hour = 0

    def is_send(self):
        if datetime.datetime.now().weekday() > 3:
            if time.localtime(time.time()).tm_hour != self.now_hour:
                self.counter = 0
            if time.localtime(time.time()).tm_hour in self.send_hour:
                self.counter += 1
                if self.counter == 1:
                    self.now_hour = time.localtime(time.time()).tm_hour
                    return True
            else:
                self.counter = 0
        else:
            self.counter = 0

        return False

    def send_class(self, col):
        wechat = WeChatRun()
        my_excel = ReadExcel('test.xlsx')
        my_data = pd.DataFrame(my_excel.get_data('开课通知'))
        # 获取需要发送的群列表,和开课通知
        for group in my_data[col].tolist():
            wechat.get_search(group)
            wechat.send_message(my_data[4][0])
            logging.info("%s: 开课通知发送成功" % group)

    def send_cost(self, weekday):
        wechat = WeChatRun()
        my_excel = ReadExcel('test.xlsx')
        my_data = pd.DataFrame(my_excel.get_data('耗课通知'), columns=['剩余课', '姓名', '昵称', '已上', '已购', '群名字', '所属班级', '通知时间', '是否通知'])
        # 获取需要通知的所有数据
        last_data = my_excel.get_row(my_excel.get_row(my_data, '是', '是否通知'), weekday, '通知时间')
        sheet = my_excel.book['耗课通知']
        print(my_data)
        for data in last_data.values:
            wechat.get_search(data[5])
            # 被通知的用户，设置自动减1.
            msg = "%s家长，上周核对%s的剩余课时数为: %s，扣除本周耗课，%s目前剩余课时数为: %s。请知悉" % (data[2], data[2], data[0], data[2], data[0] - 1)
            print(msg)
            # 剩余课时数减1 ，通过openpyxl 修改sheet中神域课时列对应的课时值
            sheet.cell(my_data[my_data['姓名'] == data[1]].index.values[0] + 2, 1).value -= 1
            my_excel.book.save('test.xlsx')
            wechat.send_message(msg)
            logging.info("%s: 耗课通知发送成功" % data[5])


def main():
    my_rules = MyRules()
    rate = 0
    while True:
        # 周五
        if datetime.datetime.now().weekday() == 4:

            if my_rules.is_send():
                if time.localtime(time.time()).tm_hour == 9:
                    my_rules.send_class(col=1)
                if time.localtime(time.time()).tm_hour == 21:
                    my_rules.send_cost(weekday='周五21:00')

        # 周六
        if datetime.datetime.now().weekday() == 5:
            # 获取周五需要发送的群列表,和开课通知
            if my_rules.is_send():
                if time.localtime(time.time()).tm_hour == 9:
                    my_rules.send_class(col=2)
                if time.localtime(time.time()).tm_hour == 13:
                    my_rules.send_cost(weekday='周六13:00')
                if time.localtime(time.time()).tm_hour == 16:
                    my_rules.send_cost(weekday='周六16:00')
                if time.localtime(time.time()).tm_hour == 18:
                    my_rules.send_cost(weekday='周六18:00')
        # 周天
        if datetime.datetime.now().weekday() == 6:
            # 获取周六需要发送的群列表,和开课通知
            if my_rules.is_send():
                if time.localtime(time.time()).tm_hour == 9:
                    my_rules.send_class(col=3)
                if time.localtime(time.time()).tm_hour == 13:
                    my_rules.send_cost(weekday='周日13:00')
                if time.localtime(time.time()).tm_hour == 16:
                    my_rules.send_cost(weekday='周日16:00')
                if time.localtime(time.time()).tm_hour == 18:
                    my_rules.send_cost(weekday='周日18:00')

        time.sleep(10)
        rate += 10
        if rate > 300:
            logging.info('pong pong pong ~')
            rate = 0

if __name__ == '__main__':

    main()
    # my_rules = MyRules()
    # my_rules.send_cost('周五21:00')
    # re = ReadExcel('test.xlsx')
    # re.delete_sheet('耗课通知')