import fitz
from main import *


# 弹窗类（添加通知弹窗）
class CollectNotice(tk.Toplevel):

    def __init__(self, file_name):
        super().__init__()
        self.title("通知")
        self.geometry("400x60+700+400")
        # 锁死窗口大小
        self.minsize(400, 60)  # 最小尺寸
        self.maxsize(400, 60)  # 最大尺寸
        self.file_name = file_name

        # 弹窗界面
        # 第一行
        self.row1 = tk.Frame(self)
        self.row1.pack(fill="x")
        tk.Label(self.row1, text="正在提取 {}.pdf 中的图片，请稍等！".format(self.file_name), font=('微软雅黑', 9), width=50).pack()

    def success(self):
        tk.Label(self.row1, text="{}.pdf 图片提取成功！".format(self.file_name), font=('微软雅黑', 9), width=50).pack()
        time.sleep(2)

    def close(self):
        self.destroy()


def collect_img(pdf_path, img_path):
    """
    pdf提取图片
    :param pdfPath: pdf文件路径
    :param img_path: 提取图片的保存路径
    """

    # 查找文件名
    search_obj = re.search('[^<>/\\\|:""\*\?]+\.', pdf_path)
    file_name = search_obj.group()[:-1]

    new_item = CollectNotice(file_name)
    doc = fitz.open(pdf_path)
    imgcount = 0
    for page in doc:
        imageList = page.getImageList()
        for imginfo in imageList:
            if imginfo:
                pix = fitz.Pixmap(doc, imginfo[0])
                img_path = img_path + '\\{}'.format(file_name) + '_{}.png'.format(imgcount)
                print(img_path)
                pix.writePNG(os.path.join(img_path))
            imgcount += 1
    new_item.success()
    new_item.close()


def pdf_collect_img(main_fun):
    # 保持转换功能同时只会运行一次
    if threading.active_count() > 2:
        stop_thread(threading.enumerate()[-1])

    # 隐藏主界面
    main_fun.withdraw()

    root = tk.Tk()
    root.title("pdf提取图片小工具 1.0版")
    root.geometry("600x300+500+200")

    # 锁死窗口大小
    root.minsize(600, 300)  # 最小尺寸
    root.maxsize(600, 300)  # 最大尺寸

    # Frame
    fm1 = tk.Frame(root, bg='black')
    fm1.titleLabel = tk.Label(fm1, text="pdf提取图片小工具 版本号: 1.0 ", font=('微软雅黑', 15), fg="white", bg='black')
    fm1.titleLabel.pack()
    fm1.pack(side=tk.TOP, fill='x', pady=5)

    tkinter.messagebox.showinfo('使用说明！必看！', '请先选择要提取图片的pdf文件(支持批量选择)，\n然后选择提取图片的存储地址。\n最后点击开始提取即可。')

    # Row
    row1 = tk.Frame(root)
    row1.pack(fill="x")
    row2 = tk.Frame(root)
    row2.pack(fill="x")
    row3 = tk.Frame(root)
    row3.pack(fill="x")

    # 选择word文件的位置
    set_pdf_button = tk.Button(row1, text="选择pdf文件", command=lambda: thread_it(set_pdf_path), font=('微软雅黑', 12),
                                width=17, height=1, bg="yellow")
    set_pdf_button.grid(row=0, column=0, padx=10, pady=20)

    set_pdf_text = tk.Text(row1, font=('微软雅黑', 10), width=45, height=2)
    set_pdf_text.grid(row=0, column=1, ipady=5)

    # 选择pdf文件存储地址
    set_img_button = tk.Button(row2, text="选择图片存储地址", command=lambda: thread_it(set_img_path), font=('微软雅黑', 12),
                               width=17, height=1, bg="yellow")
    set_img_button.grid(row=0, column=0, padx=10, pady=20)

    set_img_text = tk.Text(row2, font=('微软雅黑', 12), width=40, height=1)
    set_img_text.grid(row=0, column=1, ipady=5)

    # 转换按钮
    collect_button = tk.Button(row3, text="开始提取", command=lambda: thread_it(collect), font=('微软雅黑', 12),
                               width=20, height=1, bg="yellow")
    collect_button.grid(row=0, column=0, padx=48, pady=20)

    # 返回按钮
    back_button = tk.Button(row3, text="返回", command=lambda: thread_it(back), font=('微软雅黑', 12),
                               width=20, height=1, bg="yellow")
    back_button.grid(row=0, column=1, padx=48, pady=20)

    # 右上角窗口退出按钮 清理内存
    root.protocol("WM_DELETE_WINDOW", lambda: sys.exit(0))

    # 设置word文件的路径函数
    def set_pdf_path():
        global pdf_path
        pdf_path = tkinter.filedialog.askopenfilenames(filetypes=[("pdf", ".pdf")], initialdir='C:/users/~/desktop',
                                                       initialfile='*.pdf', title='选择pdf文件')

        # 计数器
        i = 1

        set_pdf_text = tk.Text(row1, font=('微软雅黑', 10), width=45, height=2)
        set_pdf_text.grid(row=0, column=1, ipady=5)

        if pdf_path:
            # print(pdf_path.name)
            for path in pdf_path:
                set_pdf_text.config(state=tk.NORMAL)
                set_pdf_text.insert(tk.CURRENT, path)
                set_pdf_text.insert(tk.CURRENT, '\n')
                i += 1

            set_pdf_text.delete(float(i), tk.END)
            set_pdf_text.config(state=tk.DISABLED, highlightbackground='black')
            tkinter.messagebox.showinfo('通知', 'pdf文件选择成功，请选择提取的img文件存储地址！')
        else:
            tkinter.messagebox.showwarning('警告', 'pdf文件选择失败，请重新选择！')

    # 设置word文件的存储路径函数
    def set_img_path():
        global img_path
        img_path = tkinter.filedialog.askdirectory(initialdir='C:/users/~/desktop', mustexist=True,
                                                   title='保存img文件')

        set_img_text = tk.Text(row2, font=('微软雅黑', 12), width=40, height=1)
        set_img_text.grid(row=0, column=1, ipady=5)

        if img_path:
            set_img_text.config(state=tk.NORMAL)
            set_img_text.insert(tk.CURRENT, img_path)
            set_img_text.config(state=tk.DISABLED, highlightbackground='black')
            tkinter.messagebox.showinfo('通知', 'pdf文件存储地址选择成功，请点击转换按钮开始转换！')
        else:
            tkinter.messagebox.showwarning('警告', 'pdf文件存储地址选择失败，请重新选择！')

    # 转换功能函数
    def collect():
        # 保持转换功能同时只会运行一次
        if threading.active_count() > 2:
            stop_thread(threading.enumerate()[-1])

        if pdf_path and img_path:
            for pdf in pdf_path:
                try:
                    print(pdf, img_path)
                    collect_img(pdf, img_path)
                except pywintypes.com_error:
                    tkinter.messagebox.showwarning('警告', 'img文件已存在，请重新选择路径！')
            tkinter.messagebox.showinfo('通知', 'pdf图片提取任务完成！')

        else:
            tkinter.messagebox.showwarning('警告', 'pdf文件或img文件存储地址选择失败，请重新选择！')

    # 返回主界面
    def back():
        # make root visible again
        main_fun.iconify()
        main_fun.deiconify()
        root.destroy()

    root.mainloop()
