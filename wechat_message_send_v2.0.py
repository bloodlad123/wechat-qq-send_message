"""
作者：Mr.Feng
功能：定时发送wechat消息,需要先打开聊天窗口
版本：V1.0
日期：2019/07/18
"""
# coding=utf-8
import requests
from threading import Timer
import win32gui
import win32con
import win32api
import time
from pynput.mouse import Button, Controller as mController
from pynput.keyboard import Key, Controller as kController


def get_news():
    """获取金山词霸每日一句，英文和翻译"""
    url = r"http://open.iciba.com/dsapi/"
    r = requests.get(url)
    content = r.json()['content']
    note = r.json()['note']
    return content + '\n' + note


def find_window(classname,titlename):
    """使用按键精灵找到句柄"""
    handle = win32gui.FindWindow(classname,titlename)
    if handle != 0:
        win32gui.ShowWindow(handle, win32con.SW_SHOWNORMAL)  # 将窗口显示出来
        win32gui.SetForegroundWindow(handle)  # 将窗口调到前端
        left, top, right, bottom = win32gui.GetWindowRect(handle)
        print({'handle': handle, 'left': left, 'top': top, 'right': right, 'bottom': bottom})
        return {'handle': handle, 'left': left, 'top': top, 'right': right, 'bottom': bottom}
    else:
        print("找不到[%s]这个人,请尝试打开独立的聊天窗口并核对窗口名称！" % titlename)
        return False


def click_send(msg):
    """模拟鼠标点击"""
    mouse = mController()
    keyboard = kController()
    mouse.press(Button.left)
    mouse.release(Button.left)
    keyboard.type(msg) # 聊天窗口在英文状态下输出文字
    # 发送消息的快捷键是 Alt+s
    with keyboard.pressed(Key.alt):
        keyboard.press('s')
        keyboard.release('s')


def main():
    # 发送的消息内容
    msg = get_news()
    print(msg)
    classname = "ChatWnd"
    titlename = "晓涵"
    try:
        wechat_win = find_window(classname,titlename)
        if wechat_win:
            time.sleep(0.01)
            input_pos = [wechat_win['right']-50,wechat_win['bottom']-50]
            win32api.SetCursorPos(input_pos)  # 定位鼠标到输入位置
            time.sleep(1)  # 等待20秒发出
            click_send(msg)
        # t = Timer(10, main)  # 每隔5秒发送一次
        # t.start()
    except Exception as e:
        print("发送给[%s]的消息失败" % titlename, e)


if __name__ == '__main__':
    main()
