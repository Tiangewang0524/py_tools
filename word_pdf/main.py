import os
import time
import re
import win32gui
import win32api
import win32con
import ctypes
import inspect
import sys
import tkinter.messagebox
import threading
import tkinter as tk
import tkinter.filedialog
from tkinter import ttk
import pythoncom
from win32com.client import constants, gencache
import pywintypes


# kill线程相关
def _async_raise(tid, exctype):
    """raises the exception, performs cleanup if needed"""
    tid = ctypes.c_long(tid)
    if not inspect.isclass(exctype):
        exctype = type(exctype)
    res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
    if res == 0:
        raise ValueError("invalid thread id")
    elif res != 1:
        # """if it returns a number greater than one, you're in trouble,
        # and you should call it again with exc=NULL to revert the effect"""
        ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
        raise SystemError("PyThreadState_SetAsyncExc failed")


def stop_thread(thread):
    if thread.name != 'MainThread':
        _async_raise(thread.ident, SystemExit)


def stop_thread_step():
    # print(threading.active_count())
    stop_thread(threading.enumerate()[-1])


def thread_it(func, *args):
    '''将函数打包进线程'''
    def gothread(*args):
        func(*args)

    # 创建
    t = threading.Thread(target=gothread, args=args)
    # 守护 !!!
    t.setDaemon(True)
    # 启动
    t.start()


# 主程序
def main_function():

    from word_to_pdf import word_to_pdf
    from pdf_collect_img import pdf_collect_img

    word_path, pdf_path, img_path = '', '', ''

    root = tk.Tk()
    root.title("word&pdf小工具 1.0版")
    root.geometry("400x200+500+200")

    # 锁死窗口大小
    root.minsize(400, 200)  # 最小尺寸
    root.maxsize(400, 200)  # 最大尺寸

    # Frame
    fm1 = tk.Frame(root, bg='black')
    fm1.titleLabel = tk.Label(fm1, text="word&pdf小工具 版本号: 1.0 ", font=('微软雅黑', 15), fg="white", bg='black')
    fm1.titleLabel.pack()
    fm1.pack(side=tk.TOP, fill='x', pady=5)

    # Row
    row1 = tk.Frame(root)
    row1.pack(fill="x")

    # functions 按钮
    word_to_pdf_button = tk.Button(row1, text="Word 转 PDF", command=lambda: thread_it(word_to_pdf(root)),
                                   font=('微软雅黑', 12), width=17, height=1, bg="yellow")
    word_to_pdf_button.grid(row=0, column=0, padx=10, pady=20)

    pdf_collect_img_button = tk.Button(row1, text="PDF提取图片", command=lambda: thread_it(pdf_collect_img(root)),
                                       font=('微软雅黑', 12), width=17, height=1, bg="yellow")
    pdf_collect_img_button.grid(row=0, column=1, padx=10, pady=20)

    # 右上角窗口退出按钮 清理内存
    root.protocol("WM_DELETE_WINDOW", lambda: sys.exit(0))

    root.mainloop()


if __name__ == "__main__":
    main_function()
