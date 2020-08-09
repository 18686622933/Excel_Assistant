#!/usr/bin/env python3
# -*- coding:utf-8 -*-


from tkGUI.base_class import *


class About(BaseButton):
    """帮助"""

    def __init__(self, master):
        BaseButton.__init__(self, master)
        self.label = tk.Label(self.base_frame, text="版本 1.0", font=('宋体', 16))
        self.label.pack(fill='both', expand=1)


class Questions(BaseButton):
    """问题反馈"""

    def __init__(self, master):
        BaseButton.__init__(self, master)
        self.label = tk.Label(self.base_frame, text="Wechat:cbowenyy\n\nEmail:cbowen-yy@163.com", font=('宋体', 16))
        self.label.pack(fill='both', expand=1)


class Tip(BaseButton):
    """打赏"""

    def __init__(self, master):
        BaseButton.__init__(self, master)
        self.img = tk.PhotoImage(file="tk/both.png")
        self.tip_label = tk.Label(self.base_frame, image=self.img)
        self.tip_label.pack(fill='both', expand=1)
