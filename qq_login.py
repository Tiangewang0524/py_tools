import os
import time
import win32gui
import win32api
import win32con
from pykeyboard import PyKeyboard
import ctypes
import inspect
import tkinter.messagebox
import threading
import tkinter as tk
import tkinter.filedialog
from tkinter import ttk


# 弹窗类（添加qq弹窗）
class AddQQ(tk.Toplevel):
    count = 0

    def __init__(self):
        super().__init__()
        self.title("添加qq")
        self.geometry("300x70")
        # 锁死窗口大小
        self.minsize(300, 70)  # 最小尺寸
        self.maxsize(300, 70)  # 最大尺寸
        AddQQ.count += 1

        # 弹窗界面

        # 第一行
        row1 = tk.Frame(self)
        row1.pack(fill="x")
        tk.Label(row1, text="请输入要添加的qq号:", width=20).pack(side=tk.LEFT)
        self.qq = tk.StringVar()
        tk.Entry(row1, textvariable=self.qq, width=20).pack(side=tk.LEFT)

        # 第二行
        row2 = tk.Frame(self)
        row2.pack(fill="x")
        tk.Label(row2, text="请输入要添加的密码:", width=20).pack(side=tk.LEFT)
        self.pwd = tk.StringVar()
        tk.Entry(row2, textvariable=self.pwd, width=20).pack(side=tk.LEFT)

        # 第三行
        row3 = tk.Frame(self)
        row3.pack(fill="x")
        tk.Button(row3, text="确定", command=self.close).pack(side=tk.RIGHT)

    def close(self):
        self.destroy()

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


# 单个qq登录警告步骤函数
def single_qq(qq, pwd):
    if threading.active_count() > 2:
        stop_thread(threading.enumerate()[-1])
    if qq != '' and pwd != '':
        tkinter.messagebox.showwarning('警告', 'QQ批量登录器运行中，请勿移动鼠标键盘！')
        qq_window(qq, pwd)
    else:
        tkinter.messagebox.showwarning('警告', '请选择一个qq！')


# QQ登录函数
def qq_window(qq, pwd):
    global qq_path
    # 运行QQ
    os.system(r'"%s"' %(qq_path))
    time.sleep(3)
    # 获取QQ的窗口句柄，参数1是类名,参数2是QQ软件的标题
    a = win32gui.FindWindow(None, "QQ")
    # 获取QQ登录窗口的位置
    # loginid[4]参数分别包括最小化最大化的坐标
    loginid = win32gui.GetWindowPlacement(a)

    # 定义一个键盘对象
    k = PyKeyboard()

    # 把鼠标放置到登陆框的输入处, 本机分辨率（3840*2160），具体坐标自己再调整
    ctypes.windll.user32.SetCursorPos(loginid[4][0] + 250, loginid[4][1] + 250)

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


# 读取txt保存的QQ号和密码
def read_file(file_path):
    # 读取账号密码文件例如： 1234567----she123456 #账号密码四个-分隔，此文件可以自定义，但是路径要对.

    # 打开账号密码文件，并传递给其他函数。
    with open(file_path, "r") as f1:
        qq_set = f1.readlines()

    return qq_set


# 登录QQ
def qq_login(qq_set):
    if threading.active_count() > 2:
        stop_thread(threading.enumerate()[-1])
    tkinter.messagebox.showwarning('警告', 'QQ批量登录器运行中，请勿移动鼠标键盘！')
    time.sleep(1)
    # 分隔账号密码，传两个参数:[0]账号, [1]密码
    for i in qq_set:
        tx = i.split("----")
        qq_window(tx[0], tx[1])


if __name__ == "__main__":
    global select_qq, select_pwd, file_path, qq_path
    select_qq = select_pwd = file_path = qq_path = ''

    root = tk.Tk()
    root.title("阿里鸽鸽QQ批量登录器 1.0版")

    # 锁死窗口大小
    root.minsize(400, 100)  # 最小尺寸
    root.maxsize(400, 100)  # 最大尺寸

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


    tkinter.messagebox.showinfo('使用说明！必看！', '首次使用请设置好QQ客户端的位置，\n然后读取保存QQ号-密码对的txt文件。\n'
                                            '单击可以选择单个QQ登录，\n双击可以删除该QQ号-密码对。')

    # Frame
    fm1 = tk.Frame(root, bg='black')
    fm1.titleLabel = tk.Label(fm1, text="阿里鸽鸽QQ批量登录器 版本号: 1.0 ", font=('微软雅黑', 15), fg="white", bg='black')
    fm1.titleLabel.pack()
    fm1.pack(side=tk.TOP, fill='x', pady=5)

    # Row
    row1 = tk.Frame(root)
    row1.pack(fill="x")
    row2 = tk.Frame(root)
    row2.pack(fill="x")
    row3 = tk.Frame(root)
    row3.pack(fill="x")

    # 设置qq客户端的位置
    qq_button = tk.Button(row1, text="设置QQ客户端位置", command=lambda: thread_it(set_qq_path), font=('微软雅黑', 12),
                            width=17, height=1, bg="yellow")
    qq_button.grid(row=0, column=0, padx=10, pady=1)
    # 设置打开保存QQ号-密码对的txt文件的按钮
    file_button = tk.Button(row1, text="读取txt", command=lambda: thread_it(open_file), font=('微软雅黑', 12),
                        width=17, height=1, bg="yellow")
    file_button.grid(row=0, column=1, padx=10, pady=1)

    # 设置qq客户端函数
    def set_qq_path():
        global qq_path
        tkinter.messagebox.showinfo('通知', 'QQ客户端的路径默认保存在C盘或D盘的\n/Program Files/QQ/Bin/QQScLauncher.exe')
        qq_path = tkinter.filedialog.askopenfile(filetypes=[("EXE", ".exe")], initialdir='D:/Program Files/QQ/Bin/',
                                                 initialfile='D:/Program Files/QQ/Bin/QQScLauncher.exe',
                                                 title='默认为C盘或D盘的/Program Files/QQ/Bin/QQScLauncher.exe')
        # print(qq_path.name)
        if qq_path:
            qq_path = qq_path.name
            # print(qq_path[-16:])
            if qq_path[-16:] == 'QQScLauncher.exe':
                tkinter.messagebox.showinfo('通知', 'QQ客户端设置成功，请点击读取txt！')
            else:
                tkinter.messagebox.showwarning('警告', 'QQ客户端设置错误，请重新选取！')
                qq_path = ''
        else:
            tkinter.messagebox.showwarning('警告', 'QQ客户端尚未设置，请重新选取！')
            qq_path = ''

    # 打开txt文件函数
    def open_file():
        global file_path, qq_path
        file_path = tkinter.filedialog.askopenfilename(filetypes=[("TXT", ".txt")])
        # 测试txt文件的格式是否正确
        with open(file_path, "r") as f1:
            try:
                first_row = f1.readline()
                test = first_row.split("----")
                test_answer = 1
                if len(test) != 2:
                    test_answer = 0
            except UnicodeDecodeError:
                tkinter.messagebox.showwarning('警告', 'txt文件内容格式错误，请重新选取！')
                file_path = ''

        # print(file_path)
        if qq_path != '':
            if file_path != '':
                if test_answer:
                    tkinter.messagebox.showinfo('通知', '点击读取txt成功！')
                    main_function()
                else:
                    tkinter.messagebox.showwarning('警告', 'txt文件内容格式错误，请重新选取！')
                    file_path = ''
            else:
                tkinter.messagebox.showwarning('警告', 'txt文件读取失败，请重新选取！')
                file_path = ''
        else:
            tkinter.messagebox.showwarning('警告', 'QQ客户端设置失败或尚未设置，请重新选取！')
            qq_path = ''

    # 主要功能函数
    def main_function():

        if file_path:
            # QQ号-密码对
            qq_set = read_file(file_path)

            # 锁死窗口大小
            root.minsize(400, 500)  # 最小尺寸
            root.maxsize(400, 500)  # 最大尺寸

            # Button
            button1 = tk.Button(row1, text="单个登录", command=lambda: thread_it(single_qq, select_qq, select_pwd), font=('微软雅黑', 12),
                                width=17, height=1, bg="yellow")
            button1.grid(row=0, column=0, padx=10, pady=1)

            button2 = tk.Button(row1, text="批量登录", command=lambda: thread_it(qq_login, qq_set), font=('微软雅黑', 12),
                                width=17, height=1, bg="yellow")
            button2.grid(row=0, column=1, padx=10, pady=1)

            button3 = tk.Button(row3, text="添加QQ号", command=lambda: thread_it(add_qq), font=('微软雅黑', 12),
                                width=17, height=1, bg="yellow")
            button3.grid(row=0, column=0, padx=10, pady=1)

            button4 = tk.Button(row3, text="保存txt", command=lambda: thread_it(save_qqset), font=('微软雅黑', 12),
                                width=17, height=1, bg="yellow")
            button4.grid(row=0, column=1, padx=10, pady=1)

            # Table
            table = ttk.Treeview(row2, show="headings")
            table["columns"] = ("QQ号", "密码")
            table.column("QQ号", width=180, anchor='center')
            table.column("密码", width=180, anchor='center')
            table.heading("QQ号", text="QQ号")
            table.heading("密码", text="密码")
            table.grid(row=0, column=0, padx=20, pady=10, ipady=60)

            # 给table添加数据
            for i in qq_set:
                row = 0
                tx = i.split("----")
                table.insert("", row, values=(tx[0], tx[1]))
                row += 1

            # 表格单击事件（单击选择一个条目进行单个qq登录）
            def table_click(event):
                global select_qq, select_pwd
                for item in table.selection():
                    content = table.item(item, "values")
                    # print(content)
                    select_qq = content[0]
                    select_pwd = content[1]

            # 表格双击事件（双击一个条目即可删除它）
            def table_double_click(event):
                for item in table.selection():
                    table.delete(item)
                    tkinter.messagebox.showinfo('通知', 'QQ删除成功！')

            table.bind('<ButtonRelease-1>', table_click)
            table.bind('<Double-Button-1>', table_double_click)

            # 滚动条
            m_scrl = tk.Scrollbar(row2, width=15)
            m_scrl.grid(row=0, column=0, padx=21, ipady=120, sticky=tk.E)
            table.configure(yscrollcommand=m_scrl.set)
            m_scrl['command'] = table.yview

            # 添加新的QQ号进登录器
            def add_qq():
                if AddQQ.count < 1:
                    # 如果误操作开启多个弹窗，则不会开启多个弹窗
                    new_item = AddQQ()
                    root.wait_window(new_item)  # 这一句很重要！
                    new_qq = new_item.qq.get()
                    new_pwd = new_item.pwd.get()
                    if new_qq != '' and new_pwd != '':
                        if 5 <= len(new_qq) <= 10:
                            table.insert("", len(table.get_children()), values=(new_qq, new_pwd))
                            tkinter.messagebox.showinfo('通知', '新QQ添加成功！')
                        else:
                            tkinter.messagebox.showwarning('警告', 'QQ号位数不正确！')
                    else:
                        tkinter.messagebox.showwarning('警告', 'QQ号/密码不能为空！')
                    AddQQ.count = 0

            # 保存/更新新的QQ密码对文件
            def save_qqset():
                save_file = tkinter.filedialog.asksaveasfile(filetypes=[("TXT", ".txt")], initialdir='C:/users',
                                                             initialfile='qq_password.txt',
                                                             title='保存txt文件')
                # print(save_file)
                if save_file:
                    save_file = save_file.name
                    new_qq_set = table.get_children()
                    with open(save_file, 'w+') as f2:
                        for item in new_qq_set:
                            row_set = table.item(item, "values")
                            row_set = '----'.join('%s' %num for num in row_set)
                            # 添加换行
                            if row_set[-1] != '\n':
                                row_set += '\n'
                            f2.write(row_set)
                    tkinter.messagebox.showinfo('通知', 'txt保存成功，保存在%s' % save_file)
                else:
                    tkinter.messagebox.showwarning('警告', 'txt路径尚未选取！')

    root.mainloop()
