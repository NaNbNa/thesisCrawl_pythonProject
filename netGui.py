import tkinter as tk
from tkinter import *
from tkinter import ttk, Tk
import time  
from tkinter import messagebox  # 打开tkiner的消息提醒框
from tkinter import filedialog # 在Gui中打开文件浏览
import socket
import googleGui
import cnkiGui
from tkinter.messagebox import askyesno
from tkinter.scrolledtext import ScrolledText
import threading
import webbrowser


class MY_GUI():
     
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name
        # NetConnected
        net = self.is_connected()
        if net==False:
            self.show_network_error_dialog()
        # url
        self.get_url = StringVar()  # 设置可变内容
        self.display_url = StringVar()  # 用于显示给用户看的变量 
        # search_word
        self.get_search_word = StringVar(value="金融科学")
        # book_path
        self.get_book_path = StringVar()
        # article_num
        self.get_article_num = StringVar(value="30")
        self.dialog_var = tk.StringVar(value="") 
        # url_map
        self.url_map = {  
        "谷歌学术镜像网": "https://so2.cljtscd.com/scholar?start=",  
        "中国知网镜像网": "http://search.cnki.com.cn/Search/ListResult"  
        }
        self.crawling_thread = None
        time.sleep(1)  # 休眠1秒

        
    def set_init_window(self):
        self.init_window_name.title("文献爬取工具")
        # {宽}* {高}
        self.init_window_name.geometry(f"{910}x{740}") 
        # 设置窗口为可调整大小  
        self.init_window_name.resizable(True, True)
        # 设置组件随窗口变化大小
        self.init_window_name.columnconfigure(0, weight=1)
        self.init_window_name.rowconfigure(0, weight=1)
        # 配置
        labelframe = LabelFrame(width=800, height=1000, text="配置")  # 框架，以下对象都是对于labelframe中添加的
        labelframe.grid(column=0, row=0,padx=10, pady=10, sticky="nsew")
        # get_book_path
        self.label = Label(labelframe, text=".xlsx目录路径: ").grid(column=0, row=1)
        self.path = Entry(labelframe, width=12, textvariable=self.get_book_path).grid(column=1, row=1)
        # 路径或者目录选择
        self.file = Button(labelframe, text="添加.xlsx文件目录", command=self.add_book_path).grid(column=2, row=1)
        # 路径菜单  
        self.options = ['File', 'Directory']  
        self.write_mode = ['Overwrite', 'Append']
        self.dropdown = ttk.Combobox(labelframe, values=self.options, state="readonly")
        self.dropdown.grid(column=4,row=1)
        self.dropdown.set('Directory')
        # 模式菜单
        self.mode = ttk.Combobox(labelframe, values=self.write_mode, state="readonly")
        self.mode.grid(column=6,row=1)
        self.mode.set('Overwrite')
        # 两个下拉框,其中一个变化带动另一个变化
        self.dropdown.bind("<<ComboboxSelected>>", self.on_drop_change)
        self.mode.bind("<<ComboboxSelected>>", self.on_mode_change)

        # URL
        self.url =Label(labelframe, text="URL: ").grid(column=0, row=2)
        self.url_input = ttk.Combobox(labelframe,textvariable=self.display_url, values=["谷歌学术镜像网", 
                                                          "中国知网镜像网"]
                                                          )
        self.url_input.grid(column=1, row=2)
        self.display_url.trace_add("write",self.update_url)
        self.url_input.set("谷歌学术镜像网")  
        # search_word
        self.search_word = Label(labelframe, text="搜索内容: ").grid(column=0, row=3)
        self.search_word_input = Entry(labelframe, width=10, textvariable=self.get_search_word)
        self.search_word_input.grid(column=1, row=3,columnspan=12,sticky=W)
        # articles 
        self.articles = Label(labelframe, text="文献数量: ").grid(column=0, row=4)
        self.articles_input = Entry(labelframe, width=10, textvariable=self.get_article_num)
        self.articles_input.grid(column=1, row=4,columnspan=12,sticky=W)
        # start 这将使得按钮占据其父框架的全部空间，从而实现居中效果（因为pack默认居中）  
        # 使用lambda,在按钮点击后,将函数传入线程.不使用lambda会返回结果,没点击就会运行
        self.start= Button(labelframe, text="开始爬取", command=self.start_crawling)
        # 按钮占满一整行(12列)
        self.start.grid(column=1, row=7, columnspan=12, rowspan=1, sticky='nsew')
        self.stop_button = tk.Button(labelframe, text="停止爬取", command=self.stop_crawling)  
        self.stop_button.grid(column=1, row=8, columnspan=12, rowspan=1, sticky='nsew')
        # dialog_var  使用StringVar来动态更新Label的文本  设置Label以填充其网格单元格
        #self.dialog_var = tk.StringVar(value="")  
        self.dialog_label = ScrolledText(labelframe, wrap=tk.WORD, width=80, height=10)  
        self.dialog_label.grid(column=1, row=9, columnspan=12, rowspan=1, sticky='nsew')
        # 清空按钮
        clear_button = tk.Button(labelframe, text="清空日志", command=self.clear_text)  
        clear_button.grid(column=1, row=10, columnspan=12, rowspan=1, sticky='nsew')
        clear_list_button = tk.Button(labelframe, text="清空列表", command=self.clear_list)  
        clear_list_button.grid(column=1, row=11, columnspan=12, rowspan=1, sticky='nsew')
        # self.dialog_label = tk.Label(labelframe, textvariable=self.dialog_var, width=80, height=5, wraplength=400, justify=tk.LEFT)  
        # self.dialog_label.grid(column=0, row=6, columnspan=12, sticky='ewns') 
        # article list
        self.article_frame = LabelFrame(text="文献列表")
        # columnspan 表示LabelFrame将跨越6列
        self.article_frame.grid(column=0, row=12, columnspan=12, sticky=NSEW)
        # 定义文献树形结构与滚动条
        # 容器
        self.article_tree = ttk.Treeview(self.article_frame, show="headings", height=15,columns=("a", "b", "c", "d","e","f","g"))
        # 滚动条
        self.vbar = ttk.Scrollbar(self.article_frame, orient=VERTICAL, command=self.article_tree.yview)
        self.article_tree.configure(yscrollcommand=self.vbar.set)
        # 表格的标题
        self.article_tree.column("a", width=50, anchor="center")
        self.article_tree.column("b", width=200, anchor="center")
        self.article_tree.column("c", width=150, anchor="center")
        self.article_tree.column("d", width=60, anchor="center")
        self.article_tree.column("e", width=60, anchor="center")
        self.article_tree.column("f", width=50, anchor="center")
        self.article_tree.column("g", width=300, anchor="center")
        self.article_tree.heading("a", text="ID")
        self.article_tree.heading("b", text="标题")
        self.article_tree.heading("c", text="作者")
        self.article_tree.heading("d", text="类型")
        self.article_tree.heading("e", text="出版日期")
        self.article_tree.heading("f", text="被引量")
        self.article_tree.heading("g", text="链接")
        # 放置树形结构和滚动条
        self.article_tree.grid(row=12, column=0, sticky=NSEW)
        self.vbar.grid(row=12, column=1, sticky=NS)
        self.article_tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        # 创建并启动爬取线程
        self.stop_flag = True
        

    def is_connected(self):  
        try:  
            # 创建一个 UDP socket  
            sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)  
            # 尝试连接到 Google 的公共 DNS 服务器  
            sock.connect(("8.8.8.8", 53))  
            # 连接成功，关闭 socket  
            sock.close()  
            return True  
        except socket.error:  
            # 连接失败，返回 False  
            return False
        

    def start_crawling(self):  
        # 启动爬取线程  
        self.crawling_thread = threading.Thread(target=self.startCrawl)
        self.crawling_thread.start()
        self.stop_flag = False
        self.start.config(state=tk.DISABLED)  
        self.stop_button.config(state=tk.NORMAL)

    def stop_crawling(self):  
        # 设置停止标志为 True  
        self.stop_flag = True
        self.stop_button.config(state=tk.DISABLED)  
        self.start.config(state=tk.NORMAL)

    def thread_it(self, func, *args):
        """ 将函数打包进线程 """
        self.myThread = threading.Thread(target=func, args=args)
        self.myThread .setDaemon(True)  # 主线程退出就直接让子线程跟随退出,不论是否运行完成。
        self.myThread .start()

    def show_network_error_dialog(self):  
        # 使用 tkinter 创建一个弹窗  
        result = messagebox.askyesno("网络错误", "未检测到网络连接，您想重试吗？")  
        if result:  
            # 用户点击了“是”，重试网络连接  
            net = self.is_connected()  
            if net:  
                # 如果网络已连接，可以在这里继续程序的初始化  
                print("网络连接成功！")  
            else:  
                # 如果仍然未连接，可以再次显示错误弹窗或处理错误  
                self.show_network_error_dialog()  
        else:  
            # 用户点击了“否”，退出程序  
            self.root.destroy()  
            exit()

    # 下拉选择框更新url
    def update_url(self, *args):  
        # 当用户从Combobox中选择一个选项时，更新self.get_url的值 
        selected_display_text = self.display_url.get()  # 获取显示给用户看的文本 
        self.get_url.set(self.url_map[selected_display_text])  # 设置实际的链接值 

    # 两个下拉框关联
    def on_drop_change(self, *args):  
        # 当第一个下拉框的值改变时，更新第二个下拉框的值  
        if self.dropdown.get() == 'Directory':  
            self.mode.set('Overwrite')   
        else:  
            self.mode.set('Append') 
    def on_mode_change(self, *args):
        if self.mode.get() == 'Overwrite':  
            self.dropdown.set('Directory')   
        else:  
            self.dropdown.set('File') 

    # 选择.xlsx目录或者文件
    def add_book_path(self):  

        selection_mode = self.dropdown.get()  
        if selection_mode == "Overwrite":
            self.dropdown.set('Directory')
        else:
            self.dropdown.set('File')

        selection = self.dropdown.get()
        file_path = None
        if selection == 'File':  
            # 选择文件  
            file_path = filedialog.askopenfilename() 
        elif selection == 'Directory':  
            # 选择目录  
            file_path = filedialog.askdirectory()  
        else:  
            # 如果出现未知选项，则不执行任何操作  
            return  
        # 设置路径到 StringVar  
        if file_path:  # 检查是否有有效路径被选择  
            self.get_book_path.set(file_path)  
  
    # add_book_path用到的函数
    def browse_files(self):  
        # 弹出文件/目录选择对话框  
        filepath = filedialog.askopenfilename() or filedialog.askdirectory()  
        if filepath:  
            self.path_entry.delete(0, tk.END)  
            self.path_entry.insert(0, filepath)
    
    # 更新爬取日志(,最多5行)
    def add_text(self,new_text,show_row=5):  
            self.dialog_label.insert(tk.END, "\n"+new_text)  

    def clear_text(self):  
        self.dialog_label.delete('1.0', 'end') 

    # 显示所有文献信息到文献列表# google
    def show_article_list(self,info_list):
        info_list = sorted(info_list,key=lambda x: x[3],reverse=True)
        for index, article_info in enumerate(info_list):
            # title-0 article-1 type-2 year-3 refer-5 link-6 
            self.article_tree.insert("", 'end', values=(index + 1, article_info[0], article_info[1],article_info[2],article_info[3], article_info[5],article_info[6]))
    def clear_list(self):
        for item in self.article_tree.get_children():  # 获取所有顶级项  
            self.article_tree.delete(item)  # 删除每个项

    def on_tree_select(self,event):  
        # 获取当前选中的行  
        item = self.article_tree.focus()  
        # 获取该行的链接列的值（这里假设链接在"g"列）  
        link = self.article_tree.item(item, "values")[6]  
        if link:  
            # 使用默认浏览器打开链接  
            webbrowser.open(link)

    def show_cnki_article_list(self,info_list):
        info_list = sorted(info_list,key=lambda x: x[4],reverse=True)
        for index, article_info in enumerate(info_list):
            # title-0 article-1 type-3 year-4 refer-8 link-9 
            self.article_tree.insert("", 'end', values=(index + 1, article_info[0], article_info[1],article_info[3],article_info[4], article_info[8],article_info[9]))
    
    def startCrawl(self):
        print("线程开始")
        self.add_text("线程开始....")
        if self.stop_flag:
            while self.stop_flag:
                time.sleep(1)
        
        self.dialog_var.set("^_^ 开始爬取文献...")  
        url = str(self.get_url.get())
        search_word = str(self.get_search_word.get())
        book_path = str(self.get_book_path.get())
        article_num = int(self.get_article_num.get())
        if url is None:
            self.add_text("url is None")
            return 
        if search_word is None:
            self.add_text("搜索词 is None")
            return
        if book_path is None:
            self.add_text("文件路径 is None")
            return
        if article_num is None:
            self.add_text("文献数 is None")
            return
        
        if url == "https://so2.cljtscd.com/scholar?start=":
            googleGui.main(url,search_word,book_path,article_num,self)
        else:
            cnkiGui.main(url,search_word,book_path,article_num,self)
        print("线程结束")
        self.add_text("线程结束....")
        
         

 
def gui_start():
    init_window = Tk()
    ui = MY_GUI(init_window)
    print(ui)
    ui.set_init_window()
    init_window.mainloop()


if __name__ == "__main__":
    gui_start()