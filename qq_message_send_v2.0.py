"""
作者：冯杰
功能：定时发送qq消息,需要先打开聊天窗口
版本：V1.0
日期：2019/07/18
"""
# coding=utf-8
import requests
from threading import Timer
import win32gui
import win32con
import win32clipboard as w
import time


def get_news():
    # 获取金山词霸每日一句，英文和翻译
    url = "http://open.iciba.com/dsapi/"
    r = requests.get(url)
    content = r.json()['content']
    note = r.json()['note']
    return content + '\n' + note


def msg_send(msg,name):
    # 将测试消息复制到剪切板中
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, msg)
    w.CloseClipboard()
    # 获取窗口句柄
    handle = win32gui.FindWindow(None, name)
    # 将窗口调到前台
    win32gui.ShowWindow(handle, win32con.SW_SHOWNORMAL)
    # 填充消息
    win32gui.SendMessage(handle, 770, 0, 0)
    # 回车发送消息
    win32gui.SendMessage(handle, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    win32gui.SendMessage(handle, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def main():
    # 发送的消息内容
    msg = get_news()
    # 窗口名字（好友备注，无备注的使用昵称）
    name = '晓涵'
    try:
        time.sleep(15000)
        msg_send(msg,name)
        # t = Timer(10, main)  # 每隔5秒发送一次
        # t.start()
    except Exception as e:
        print('发送失败！', e)


if __name__ == '__main__':
    main()
