import netGui
from tkinter import *
from tkinter import  Tk


def gui_start():
    init_window = Tk()
    ui = netGui.MY_GUI(init_window)
    print(ui)
    ui.set_init_window()
    init_window.mainloop()


if __name__ == "__main__":
    gui_start()