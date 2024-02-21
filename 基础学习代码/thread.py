import sys
import time
import tkinter as tk
from tkinter.messagebox import askyesno
from tkinter.scrolledtext import ScrolledText
import threading


class TestGui(object):
    def __init__(self, init_window_name):
        self.init_window_name = init_window_name
        self.init_window_name.title("将button方法打包进线程，现实运行不卡顿测试")  # 设置窗口标题
        self.init_window_name.geometry('700x350')  # 设置窗口大小
        """ 点击右上角关闭窗体弹窗事件 """
        self.init_window_name.protocol('WM_DELETE_WINDOW', lambda: self.thread_it(self.clos_window))
        """ 组件容器创建 """
        self.log_frame = tk.Frame(self.init_window_name)  # 创建存放日志组件的容器
        self.log_frame.grid(padx=20, pady=0, row=1, column=0, sticky=tk.W)
        self.runs_button_frame = tk.Frame(self.init_window_name)  # 创建存放日志组件的容器
        self.runs_button_frame.grid(padx=20, pady=0, row=2, column=0, sticky=tk.W)
        """ 日志框 """
        self.run_log = ScrolledText(self.log_frame, font=('楷体', 13), width=69, height=17)
        self.run_log.grid(padx=20, pady=5, row=0, column=0)
        """ 操作按钮 """
        self.start_run1 = tk.Button(self.runs_button_frame, text='开始打印1', font=('行楷', 15, 'bold'), fg="white", bg="#1E90FF", width=25, command=lambda: self.thread_it(self.print1))
        self.start_run1.grid(padx=20, pady=0, row=0, column=1)
        self.start_run2 = tk.Button(self.runs_button_frame, text='开始打印2', font=('行楷', 15, 'bold'), fg="white", bg="#1E90FF", width=25, command=lambda: self.thread_it(self.print2))
        self.start_run2.grid(padx=35, pady=0, row=0, column=2)

    def thread_it(self, func, *args):
        """ 将函数打包进线程 """
        self.myThread = threading.Thread(target=func, args=args)
        self.myThread .setDaemon(True)  # 主线程退出就直接让子线程跟随退出,不论是否运行完成。
        self.myThread .start()

    def print1(self):
        for i in range(100):
            tip_content = f'第{i}次打印 - 我是小洲1'
            self.run_log_print(message=tip_content)
            time.sleep(0.1)  # 睡眠
        self.run_log_print(message='我是小洲1 - 打印完成')

    def print2(self):
        for i in range(100, 200):
            tip_content = f'第{i}次打印 - 我是小洲2'
            self.run_log_print(message=tip_content)
            time.sleep(0.05)  # 睡眠
        self.run_log_print(message='我是小洲2 - 打印完成')

    def run_log_print(self, message):
        self.run_log.config(state=tk.NORMAL)
        self.run_log.insert(tk.END, "\n" + message + "\n")
        self.run_log.see(tk.END)
        self.run_log.update()
        self.run_log.config(state=tk.DISABLED)

    def clos_window(self):
        ans = askyesno(title='小洲助手v1.1警告', message='是否确定退出程序？\n是则退出，否则继续！')
        if ans:
            self.init_window_name.destroy()
            sys.exit()
        else:
            return None


if __name__ == '__main__':
    """ 把button方法打包进线程，现实运行不卡顿 """
    """ 实例化出一个父窗口 """
    init_window = tk.Tk()
    """ tk界面置顶 """
    init_window.attributes("-topmost", 1)
    """ 创建Gui类对象 """
    test_gui = TestGui(init_window)
    """ 初始化GUi组件 """
    init_window.mainloop()