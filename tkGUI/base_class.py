#!/usr/bin/env python3
# -*- coding:utf-8 -*-

import os
import tkinter as tk
from tkinter import ttk
import tkinter.messagebox
import tkinter.filedialog as filedialog


class TipButton:
    """自定义button，无边框，有鼠标悬停提示，可摧毁"""

    def __init__(self, father, show, prompt, func):
        """
        :param father: 上级控件
        :param show: 显示名称
        :param prompt: 提示内容
        :param func: 点击执行动作
        """
        self.root = father
        self.show = show
        self.help = prompt
        self.cursor = "hand2"  # 鼠标变换
        self.relief = "flat"  # 无边框
        self.command = func
        self.position = tk.RIGHT
        self.button = tk.Button(self.root,
                                text=show,
                                font=('Microsoft YaHei', 15),
                                cursor=self.cursor,
                                relief=self.relief,
                                command=self.command)

        self.button.pack(side=self.position)

        self.button.bind("<Enter>", self.on_enter)
        self.button.bind("<Leave>", self.on_leave)

    def disable(self):
        self.button.configure(state='disable')

    def normal(self):
        self.button.configure(state='normal')

    def on_enter(self, event):
        self.button.configure(text=self.help)

    def on_leave(self, enter):
        self.button.configure(text=self.show)

    def destroy(self):
        self.button.destroy()


class TipEntry(tk.Entry):
    """继承Entry类，并且封装了三个事件处理函数"""

    def set_default_text(self, text):
        self.default_text = text
        self.insert(0, self.default_text)

    def mybind(self):
        self.bind("<FocusIn>", self.focus_in)  # 获得键盘焦点
        self.bind("<FocusOut>", self.focus_out)  # 失去键盘焦点
        self.bind("<Double-1>", self.on_double)  # 双击
        # self.bind("Return", self.on_return)  # 回车事件的函数需要从类外部传入，目前没找到好方法，暂时去掉

    def focus_in(self, en=None):
        """清除默认文字，并将字体颜色设为黑色"""
        if self.get() == self.default_text:
            self.delete('0', 'end')
            self.configure(fg="black")

    def focus_out(self, en=None):
        """失去键盘焦点时，如果没有输入信息，则还原默认文字"""
        if self.get() == "":
            self.configure(fg="gray")
            self.insert(0, self.default_text)

    def on_double(self, en=None):
        """清除全部文字"""
        self.delete('0', 'end')


class TipListbox:
    """带纵向滚动条，可多选的Listbox"""

    def __init__(self, father, data):
        self.root = father
        self.data = data
        self.all_list = data

        self.listbox = tk.Listbox(self.root, selectmode=tk.EXTENDED)  # , relief="flat"
        self.listbox.pack(fill='both', side=tk.LEFT, expand=1)

        for i in range(0, len(data)):
            self.listbox.insert(i, data[i])

        # 创建Scrollbar
        self.yscrollbar = tk.Scrollbar(self.listbox, command=self.listbox.yview)
        self.yscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=self.yscrollbar.set)

    def get_file_list(self, path):
        """返回一个list，首位为目录， 后面是该地址下的文件列表"""
        index_list = self.listbox.curselection()
        if not index_list:
            tk.messagebox.showinfo("提示", "请选择要合并的文件")
        else:
            return [path] + [self.all_list[i] for i in index_list]


class TipTreeview:
    """自定义treeview，具有双向滚动条和字段排序功能"""

    def __init__(self, father, data, height=20):
        self.root = father
        self.columns = [i for i in range(len(data[0]))]
        self.height = height

        self.treeview = ttk.Treeview(self.root, columns=self.columns, height=self.height, show="headings")
        self.treeview.pack(fill='both', side=tk.LEFT, expand=1, pady=2)  #

        """纵向滚动条"""
        self.ybar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.treeview.yview)  # 设置滚动条控件
        self.ybar.pack(side=tk.LEFT, fill='y')  # 滚动条控件与被控制的Treeview同在一个容器中，并列放置，纵向填充
        self.treeview.configure(yscrollcommand=self.ybar.set)

        """横向滚动条"""
        self.xbar = ttk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.treeview.xview)
        self.treeview.configure(xscrollcommand=self.xbar.set)
        self.xbar.pack(side=tk.BOTTOM, fill='x')

        for col in self.columns:
            """结合treeview_sort_column()函数实现点击字段名排序"""
            self.treeview.heading(col, text=col,
                                  command=lambda c=col: self.treeview_sort_column(self.treeview, c, False))

        for column in self.columns:
            """设置字段显示名称及显示格式"""
            try:
                self.treeview.heading(column, text=data[0][column])
                self.treeview.column(column, width=self.get_width(str(data[1][column])), anchor='center')
            except IndexError:
                pass

        self.write(data)

    def get_width(self, values):
        """根据第一行数据的值调整Trieeview的列宽"""
        if values.isdigit():
            field_width = len(values) * 12
        else:
            field_width = len(values) * 23
        if field_width <= 30:
            return 30
        else:
            return field_width

    def treeview_sort_column(self, tv, col, reverse):
        """Treeview点击字段排序"""
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(key=lambda t: (t[0]), reverse=reverse)

        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def clear_out(self):
        """清除全部数据"""
        items = self.treeview.get_children()
        [self.treeview.delete(item) for item in items]

    def write(self, data):
        """数据写入"""
        self.columns = [i for i in range(len(data[0]))]
        for column in self.columns:
            """设置字段显示名称及显示格式"""
            try:
                self.treeview.heading(column, text=data[0][column])
                self.treeview.column(column, width=self.get_width(str(data[1][column])), anchor='center')
            except IndexError:
                pass

        for row in range(1, len(data)):
            self.treeview.insert('', 'end', values=data[row])

    def destroy(self):
        self.treeview.destroy()


class Welcome:
    """创建欢迎界面"""

    def __init__(self, master):
        self.root = master
        self.base_frame = tk.Frame(self.root)
        self.base_frame.pack(fill='both', expand=1)
        self.welcome_label = tk.Label(self.base_frame, text="Welcome!", font=('华文彩云', 80), fg="blue")
        self.welcome_label.pack(fill='both', expand=1)

    def destroy(self):
        self.base_frame.destroy()


class BaseButton:
    """在底层frame最上边加入一个返回功能的button，横向填充"""

    def __init__(self, master):
        self.root = master
        self.base_frame = tk.Frame(self.root)
        self.base_frame.pack(fill='both', expand=1)

        self.button_frame = tk.Frame(self.base_frame)
        # self.button_frame.pack(fill="x")
        self.home_button = TipButton(self.button_frame, " <<< ", "go Home", self.to_welcome)

    def destroy(self):
        self.base_frame.destroy()

    def to_welcome(self):
        self.base_frame.destroy()
        self.base_frame = Welcome(self.root)


def get_save_path(name):
    filename = filedialog.asksaveasfilename(
        defaultextension='.xlsx',  # 默认文件的扩展名
        filetypes=[('Excel Files', '*.xlsx'),
                   # ('txt Files', '*.txt'),
                   ('All Files', '*.*')],  # 设置文件类型下拉菜单里的的选项
        initialdir=os.getcwd(),  # 对话框中默认的路径
        initialfile=name,  # 对话框中初始化显示的文件名
        # parent=self.master,                #父对话框(由哪个窗口弹出就在哪个上端)
        title="输出到："  # 弹出对话框的标题
    )
    return filename
