import os
import time
import win32gui
import win32api
import win32con
# from pymouse import PyMouse
from pykeyboard import PyKeyboard
from ctypes import *


def qq_login(qq, pwd):
    # 运行QQ
    os.system('"D:\Program Files\QQ\Bin\QQScLauncher.exe"')
    time.sleep(3)
    # 获取QQ的窗口句柄，参数1是类名,参数2是QQ软件的标题
    a = win32gui.FindWindow(None, "QQ")
    # 获取QQ登录窗口的位置
    # loginid[4]参数分别包括最小化最大化的坐标
    loginid = win32gui.GetWindowPlacement(a)

    # 定义一个键盘对象
    k = PyKeyboard()

    # 把鼠标放置到登陆框的输入处, 本机分辨率（3840*2160），具体坐标自己再调整
    windll.user32.SetCursorPos(loginid[4][0] + 250, loginid[4][1] + 250)

    # 按下鼠标再释放
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(1)

    # 输入用户名
    k.type_string(qq)
    time.sleep(0.2)

    # 按下tab，切换到输入密码的地方
    k.press_key(k.tab_key)
    k.release_key(k.tab_key)

    # 输入密码
    k.type_string(pwd)
    time.sleep(0.2)

    # 按下回车
    k.press_key(k.enter_key)
    k.release_key(k.enter_key)
    time.sleep(5)


if __name__ == "__main__":
    # 读取账号密码文件例如： 1234567----she123456 #账号密码四个- 分隔
    # 此文件可以自定义，但是路径要一定对
    qq_file = 'C:\\Users\\qq_password.txt'

    # 打开账号密码文件
    with open(qq_file, "r") as f1:
        qq_set = f1.readlines()
        #分隔账号密码，传两个参数，账号密码
        for i in qq_set:
            tx = i.split("----")
            qq_login(tx[0], tx[1])

