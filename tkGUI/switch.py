#!/usr/bin/env python3
# -*- coding:utf-8 -*-


from tkGUI.help_frame import *
from tkGUI.excel_frame import *
from tkGUI.db_frame import *
from tkGUI.base_class import *


class BaseUI:
    """生成前端基本结构和一级菜单"""

    def __init__(self, master):
        self.font = {"welcome": ('华文琥珀', 80),
                     "label": ('华文琥珀', 20),
                     "text": ('宋体', 16),
                     "button": ('Monaco', 12)}
        self.root = master
        self.root.title('Excel Assistant')
        self.width, self.heigth = 900, 500
        self.root.geometry('%dx%d+%d+%d' % (self.width, self.heigth,
                                            (self.root.winfo_screenwidth() - self.width) / 2,
                                            (self.root.winfo_screenheight() - self.heigth) / 2))

        self.base_menu = tk.Menu(self.root)
        self.root.config(menu=self.base_menu)
        self.excel_menu = tk.Menu(self.base_menu, tearoff=0)
        self.db_menu = tk.Menu(self.base_menu, tearoff=0)
        self.help_menu = tk.Menu(self.base_menu, tearoff=0)
        self.base_menu.add_cascade(label='Excel', menu=self.excel_menu)
        self.base_menu.add_cascade(label='数据库', menu=self.db_menu)
        self.base_menu.add_cascade(label='帮助', menu=self.help_menu)


class PageSwitch(BaseUI):
    """所有界面的切换"""

    def __init__(self, master):
        BaseUI.__init__(self, master)

        self.functions = {
            "excel": {"查重": self.f_check, "对比": self.f_contrast, "合并": self.f_merge},
            "db": {"转in('')": self.f_format_conversion},
            "help": {"关于": self.f_about, "问题反馈": self.f_question}}

        for k, v in self.functions["excel"].items():
            self.excel_menu.add_command(label=k, command=v)
        for k, v in self.functions["db"].items():
            self.db_menu.add_command(label=k, command=v)
        for k, v in self.functions["help"].items():
            self.help_menu.add_command(label=k, command=v)

        self.base_frame = Welcome(self.root)

    def welcome_page(self):
        self.base_frame.destroy()
        self.base_frame = Welcome(self.root)

    def f_about(self):
        self.base_frame.destroy()
        self.base_frame = About(self.root)

    def f_question(self):
        self.base_frame.destroy()
        self.base_frame = Questions(self.root)

    def f_tip(self):
        self.base_frame.destroy()
        self.base_frame = Tip(self.root)

    def f_check(self):
        self.base_frame.destroy()
        self.base_frame = Check(self.root)

    def f_contrast(self):
        self.base_frame.destroy()
        self.base_frame = Contrast(self.root)

    def f_merge(self):
        self.base_frame.destroy()
        self.base_frame = Merge(self.root)

    # def f_form_integration(self):
    #     self.base_frame.destroy()
    #     self.base_frame = FormIntegration(self.root)


    def f_format_conversion(self):
        self.base_frame.destroy()
        self.base_frame = FormatConversion(self.root)


if __name__ == '__main__':
    root = tk.Tk()
    PageSwitch(root)
    root.mainloop()
