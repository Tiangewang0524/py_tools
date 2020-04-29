# py_tools

py_tools 包含了一些自己练手的基于Python写成的小工具。

## 1. file_rename
是一个可以批量重命名文件的工具，尤其适合给网上下载的起名不规则的电视剧重命名。（无界面，需要自行编译运行）

## 2. qq_login
是一个批量自动登录QQ号的工具，`QQ客户端`版本为`2020`，测试了1024\*768, 1280\*1024，1366\*768，1920\*1080，3840\*2160等分辨率。（不支持`TIM QQ`）


需要提前准备好要登录的QQ号-密码对的txt文件，每一对QQ号-密码的保存格式`QQ号----密码`。（4个-符号作分隔符）同时支持自行添加qq号，删除qq号，另存为新的txt文件。

* ### prerequisite：
        pip install win32gui
        pip install PyKeyboard
        pip install ctypes
        pip install inspect
        pip install tkinter.messagebox
        pip install tkinter.filedialog


## 3. spider_tpp
是一个在淘票票上的爬虫，用于爬取每个城市当前正在上映的电影名称，后续可能会加入主演导演，各个电影院场次安排，座位情况，票价等。（无界面，需要自行编译运行）

* `tpp.html`是淘票票动态网页的保存版，变成静态网页方便解析。编译时需要指定该html文件的路径。


* #### 输出结果示例：
        [('阿坝', '无影片上映'), ('阿克苏', '无影片上映'), ('阿拉善', '无影片上映'), ('安庆', '无影片上映'),  ('安康', '有影片上映影片名：误杀')]
