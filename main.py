# encoding:utf-8
import os
from tkinter import *
from tkinter import filedialog
import urllib.request
import shutil
import pandas as pd
# 生成。exe命令 pyinstaller  -F -w download.py


class FileDownload:
    def __init__(self, master):
        self.filename = StringVar()
        self.filename.set('选择要打开的文件信息')
        self.src_dir = ''
        self.urls = []
        self.name_url_dict = {}
        # tkinter刚开始不熟悉，布局麻烦，这里使用的网格布局，row，column 的index，定位cell
        open_btn = Button(master, text="打开文件", width=15, height=3, borderwidth=2,
                          command=self.openfile)
        download_btn = Button(master, text="开始下载", width=15, height=3, borderwidth=2,
                              command=self.download)
        open_btn.grid(row=0, column=0, ipadx=5, ipady=3, padx=10, pady=3)
        download_btn.grid(row=1, column=0, ipadx=5, ipady=3, padx=10, pady=3)
        self.file_label = Label(master, textvariable=self.filename, wraplength=250, fg='green', bg='white', )
        self.file_label.grid(row=0, column=1, rowspan=2, ipadx=200, ipady=15, padx=10, pady=2)  # ,padx=20,pady=10,
        self.text = Text(master, width=100, height=50, bg='grey', fg='blue')
        self.text.grid(row=2, columnspan=2, ipadx=10, ipady=10)

    def openfile(self):
        # 系统默认打开用户的桌面文件
        default_dir = os.path.join(os.path.expanduser("~"), 'Desktop')
        fname = filedialog.askopenfilename(title='选择打开的文件', filetypes=[('All Files', '*')],
                                           initialdir=(os.path.expanduser(default_dir)))
        print('打开的文件', fname)
        self.src_dir = fname[0:fname.rindex('/')]
        self.filename.set(fname)
        self.urls = self.read_urls(fname)

        self.text.insert(END, '文件加载完毕，请点击下载按钮\n')
        self.text.see(END)
        self.text.update()  # 这个很关键，更新页面text信息



    # 读取urls
    def read_urls(self, fname):
        urls = []
        df = pd.read_excel(fname, sheet_name="file")
        for idx, row in df.iterrows():
            name = row["name"]
            url = row["url"].replace("\n", "|")
            urls.append(name + "\t" + url)
        return urls

    def print_schedule(self, a, b, c):
        '''
        打印进度条信息
        a:已经下载的数据块
        b:数据块的大小
        c:远程文件的大小
        '''
        per = 100.0 * a * b / c
        if per > 100:
            per = 100
        print('下载进度%.1f%%' % per)

    def download(self):
        # 处理文件中的重复和多下载链接

        # 创建文件下载目录
        file_dir = os.path.join(self.src_dir, "reslut")
        if os.path.isdir(file_dir):
            shutil.rmtree(file_dir)
            os.makedirs(file_dir)
        else:
            os.makedirs(file_dir)
        # 处理同一个名字多个url链接和同一个名字多次提交
        for url in self.urls:
            name = url.split('\t')[0]
            name_url = url.split('\t')[1]
            if self.name_url_dict.get(name) is None:
                self.name_url_dict[name] = name_url
            else:
                url_temp = self.name_url_dict.get(name)
                self.name_url_dict[name] = name_url + '|' + url_temp


        # 文件下载
        i = 1
        urls_nums = len(self.name_url_dict)
        for (name, name_url) in self.name_url_dict.items():
            self.text.insert(END, '正在下载%d/%d\n' % (i, urls_nums))
            self.text.see(END)
            self.text.update()  # 这个很关键，更新页面text信息
            if name_url.find("|") == -1:
                file_type = name_url.split('type=')[1]
                if len(name) > 20:
                    name = name[-20:]
                local = os.path.join(file_dir, name)
                urllib.request.urlretrieve(name_url, local + "." + file_type, self.print_schedule)
            else:
                os.makedirs(file_dir + "/" + name)
                index = 1
                for url in name_url.split("|"):
                    file_type = url.split('type=')[1]
                    if len(name) > 20:
                        name = name[-20:]
                    local = os.path.join(file_dir, name, name)
                    urllib.request.urlretrieve(url, local + str(index) + "." + file_type, self.print_schedule)
                    index += 1
            i += 1
        self.text.insert(END, '下载结束')
        self.text.see(END)
        self.text.update()


if __name__ == '__main__':
    tker = Tk()
    tker.title("自动文件下载器")
    tker.columnconfigure(0, weight=3)
    tker.columnconfigure(1, weight=7)
    tker.rowconfigure(0, weight=1)
    tker.rowconfigure(1, weight=1)
    tker.rowconfigure(2, weight=8)
    tker.geometry('640x480')  # 设置主窗口的初始大小640x480
    app = FileDownload(tker)
    tker.mainloop()
