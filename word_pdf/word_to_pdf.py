from main import *


# 弹窗类（添加通知弹窗）
class ConvertNotice(tk.Toplevel):

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
        tk.Label(self.row1, text="正在将 {}.doc 转换为PDF中，请稍等！".format(self.file_name), font=('微软雅黑', 9), width=50).pack()

    def success(self):
        tk.Label(self.row1, text="{}.doc 转换成功！".format(self.file_name), font=('微软雅黑', 9), width=50).pack()
        time.sleep(2)

    def close(self):
        self.destroy()


# word 转 pdf
def createPdf(wordPath, pdfPath, file_name):
    """
    word转pdf
    :param wordPath: word文件路径
    :param pdfPath: 生成pdf文件路径
    :param file_name: word文件名
    """
    # tkinter.messagebox.showinfo('通知', 'Word转换PDF中，点击确定继续！')
    new_item = ConvertNotice(file_name)
    pythoncom.CoInitialize()
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(wordPath, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfPath,
                            constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    new_item.success()
    new_item.close()
    word.Quit(constants.wdDoNotSaveChanges)

def word_to_pdf(main_fun):
    # 保持转换功能同时只会运行一次
    if threading.active_count() > 2:
        stop_thread(threading.enumerate()[-1])

    # 隐藏主界面
    main_fun.withdraw()

    root = tk.Tk()
    root.title("word转pdf小工具 1.0版")
    root.geometry("600x300+500+200")

    # 锁死窗口大小
    root.minsize(600, 300)  # 最小尺寸
    root.maxsize(600, 300)  # 最大尺寸

    # Frame
    fm1 = tk.Frame(root, bg='black')
    fm1.titleLabel = tk.Label(fm1, text="word转pdf小工具 版本号: 1.0 ", font=('微软雅黑', 15), fg="white", bg='black')
    fm1.titleLabel.pack()
    fm1.pack(side=tk.TOP, fill='x', pady=5)

    tkinter.messagebox.showinfo('使用说明！必看！', '请先选择要转换的word文件(支持批量选择)，\n然后选择转换生成的pdf文件存储地址。\n最后点击开始转换即可。')

    # Row
    row1 = tk.Frame(root)
    row1.pack(fill="x")
    row2 = tk.Frame(root)
    row2.pack(fill="x")
    row3 = tk.Frame(root)
    row3.pack(fill="x")

    # 选择word文件的位置
    set_word_button = tk.Button(row1, text="选择word文件", command=lambda: thread_it(set_word_path), font=('微软雅黑', 12),
                                width=17, height=1, bg="yellow")
    set_word_button.grid(row=0, column=0, padx=10, pady=20)

    set_word_text = tk.Text(row1, font=('微软雅黑', 10), width=45, height=2)
    set_word_text.grid(row=0, column=1, ipady=5)

    # 选择pdf文件存储地址
    set_pdf_button = tk.Button(row2, text="选择pdf文件存储地址", command=lambda: thread_it(set_pdf_path), font=('微软雅黑', 12),
                               width=17, height=1, bg="yellow")
    set_pdf_button.grid(row=0, column=0, padx=10, pady=20)

    set_pdf_text = tk.Text(row2, font=('微软雅黑', 12), width=40, height=1)
    set_pdf_text.grid(row=0, column=1, ipady=5)

    # 转换按钮
    convert_button = tk.Button(row3, text="开始转换", command=lambda: thread_it(convert), font=('微软雅黑', 12),
                               width=20, height=1, bg="yellow")
    convert_button.grid(row=0, column=0, padx=48, pady=20)

    # 返回按钮
    back_button = tk.Button(row3, text="返回", command=lambda: thread_it(back), font=('微软雅黑', 12),
                            width=20, height=1, bg="yellow")
    back_button.grid(row=0, column=1, padx=48, pady=20)

    # 右上角窗口退出按钮 清理内存
    root.protocol("WM_DELETE_WINDOW", lambda: sys.exit(0))

    # 设置word文件的路径函数
    def set_word_path():
        global word_path
        word_path = tkinter.filedialog.askopenfilenames(filetypes=[("doc", ".doc"), ("docx", ".docx")],
                                                        initialdir='C:/users/~/desktop',
                                                        initialfile='*.docx',
                                                        title='选择word文件')

        # 计数器
        i = 1

        set_word_text = tk.Text(row1, font=('微软雅黑', 10), width=45, height=2)
        set_word_text.grid(row=0, column=1, ipady=5)

        if word_path:
            # print(word_path.name)
            for path in word_path:
                set_word_text.config(state=tk.NORMAL)
                set_word_text.insert(tk.CURRENT, path)
                set_word_text.insert(tk.CURRENT, '\n')
                i += 1

            set_word_text.delete(float(i), tk.END)
            set_word_text.config(state=tk.DISABLED, highlightbackground='black')
            tkinter.messagebox.showinfo('通知', 'word文件选择成功，请选择生成的pdf文件存储地址！')
        else:
            tkinter.messagebox.showwarning('警告', 'word文件选择失败，请重新选择！')

    # 设置word文件的存储路径函数
    def set_pdf_path():
        global pdf_path
        pdf_path = tkinter.filedialog.askdirectory(initialdir='C:/users/~/desktop', mustexist=True,
                                                   title='保存pdf文件')
        # pdf_path = tkinter.filedialog.asksaveasfilename(initialdir='C:/users/~/desktop', filetypes=[("pdf", ".pdf")], initialfile='*.pdf',
        #                                            title='保存pdf文件')
        set_pdf_text = tk.Text(row2, font=('微软雅黑', 12), width=40, height=1)
        set_pdf_text.grid(row=0, column=1, ipady=5)

        if pdf_path:
            set_pdf_text.config(state=tk.NORMAL)
            set_pdf_text.insert(tk.CURRENT, pdf_path)
            set_pdf_text.config(state=tk.DISABLED, highlightbackground='black')
            tkinter.messagebox.showinfo('通知', 'pdf文件存储地址选择成功，请点击转换按钮开始转换！')
        else:
            tkinter.messagebox.showwarning('警告', 'pdf文件存储地址选择失败，请重新选择！')

    # 转换功能函数
    def convert():
        # 保持转换功能同时只会运行一次
        if threading.active_count() > 2:
            stop_thread(threading.enumerate()[-1])
        if word_path and pdf_path:
            for word in word_path:
                try:
                    word = word.replace('/', '\\')
                    search_obj = re.search('[^<>/\\\|:""\*\?]+\.', word)
                    file_name = search_obj.group()[:-1]
                    pdf = pdf_path + '/' + file_name + '.pdf'
                    pdf = pdf.replace('/', '\\')
                    createPdf(word, pdf, file_name)
                except pywintypes.com_error:
                    tkinter.messagebox.showwarning('警告', 'pdf文件已存在，请重新选择路径！')
            tkinter.messagebox.showinfo('通知', 'pdf文件转换任务完成！')

        else:
            tkinter.messagebox.showwarning('警告', 'word文件或pdf文件存储地址选择失败，请重新选择！')

    # 返回主界面
    def back():
        # make root visible again
        main_fun.iconify()
        main_fun.deiconify()
        root.destroy()

    root.mainloop()