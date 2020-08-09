#!/usr/bin/env python3
# -*- coding:utf-8 -*-

from tkGUI.base_class import *
import tkinter.messagebox
from functions.data_format import *


class FormatConversion:
    """excel转in('')"""

    class Page:
        """back、next键，以及中间空白容器"""

        def __init__(self, father, go_back, go_next):
            self.root = father
            self.func_frame = tk.Frame(self.root)
            self.func_frame.pack(fill='both', expand=1)

            self.left_frame = tk.Frame(self.func_frame)
            self.left_frame.pack(side="left", fill="y")
            self.right_frame = tk.Frame(self.func_frame)
            self.right_frame.pack(side="right", fill="y")

            self.center_frame = tk.Frame(self.func_frame)
            self.center_frame.pack(fill='both', expand=1)

            self.back_button = TipButton(self.left_frame, "  <   ", "back", go_back)
            self.next_button = TipButton(self.right_frame, "   >  ", "next", go_next)
            self.back_button.disable()
            self.next_button.disable()

    class Pagei:
        """第一页完整功能"""

        def __init__(self, father, next_button):
            """还Page的中间容器上添加了上、下两个容易，分别放置选择文件的button，和展示文件路径的label"""
            self.root = father
            self.next_button = next_button

            # 当前页面最底层frame
            self.iframe = tk.Frame(self.root)
            self.iframe.pack(fill="both", expand=1)

            self.up_frame = tk.Frame(self.iframe)
            self.middle_frame = tk.Frame(self.iframe)
            self.down_frame = tk.Frame(self.iframe)

            self.up_frame.pack(fill="both", expand=1)
            self.middle_frame.pack(fill="x")
            self.down_frame.pack(fill="both", expand=1)

            self.field = TipEntry(self.middle_frame, width=45, fg="gray", justify="center")  # 自定义Entry类，包含事件处理
            self.field.set_default_text("选取单元格")
            self.field.mybind()  # 生成事件
            self.field.pack()

            self.button = tk.Button(self.up_frame,
                                    text=" Select the file and make sure the sheet is first. ",
                                    font=('Microsoft YaHei', 26),
                                    cursor="hand2",
                                    relief="flat",
                                    command=self.ask_file)
            self.button.pack(side="bottom")
            self.label = tk.Label(self.down_frame, text="", font=('Microsoft YaHei', 15))
            self.label.pack()

            self.address = ""

        def ask_file(self):
            """选择文件函数"""
            address = filedialog.askopenfilename()
            base_name = os.path.basename(address)
            if not address:
                self.button.configure(text=" Select the file and make sure the sheet is first. ")
                self.label.configure(text="")
                self.next_button.disable()
            else:
                if base_name and not (base_name.endswith('.xls') or base_name.endswith('.xlsx')):
                    tk.messagebox.showwarning('提示', '请选择正确的Excel文件！')
                else:
                    self.button.configure(text=" Click to re-select ")
                    self.label.configure(text=address)
                    self.next_button.normal()
                    if len(address) >= 70:
                        self.label.configure(font=('Microsoft YaHei', 9))
                    elif len(address) >= 65:
                        self.label.configure(font=('Microsoft YaHei', 10))
                    elif len(address) >= 60:
                        self.label.configure(font=('Microsoft YaHei', 11))
                    elif len(address) >= 55:
                        self.label.configure(font=('Microsoft YaHei', 13))
                    else:
                        self.label.configure(font=('Microsoft YaHei', 15))

                    self.address = address

        def get_cells(self):
            return self.field.get()

        def pack_forget(self):
            self.iframe.pack_forget()

        def pack(self):
            self.iframe.pack(fill='both', expand=1)

    class Pageii:
        """第二页完整功能"""

        def __init__(self, father, path, cells):
            self.root = father
            self.path = path
            self.cells = cells
            self.iiframe = tk.Frame(self.root)  # 本页面此层frame
            self.iiframe.pack(fill="both", expand=1)
            self.ctrl_frame = tk.Frame(self.iiframe)  # 最上行放置控制按键的frame
            self.info_frame = tk.Frame(self.iiframe)
            self.ctrl_frame.pack(fill="x", side="top")  # 展示信息的活页
            self.info_frame.pack(fill="both", expand=1)

            self.label = tk.Label(self.ctrl_frame)

            self.result_text = tk.Text(self.info_frame, pady=1)
            self.vbar = ttk.Scrollbar(self.info_frame, orient=tk.VERTICAL, command=self.result_text.yview)  # 设置滚动条控件
            self.result_text.configure(yscrollcommand=self.vbar.set)

            self.format_data = Format(path).get_cell_list(cells)
            self.result_text.insert("end", self.format_data)

        def show_result(self):
            self.label.pack(side=tk.LEFT)
            self.result_text.pack(fill='both', expand=1, side=tk.LEFT)
            self.vbar.pack(side=tk.LEFT, fill='y')

        def change_label(self, address):
            self.label.configure(text="%s" % os.path.basename(address))

        def pack_forget(self):
            self.iiframe.pack_forget()

        def pack(self):
            self.iiframe.pack(fill='both', expand=1)

    def __init__(self, master):
        self.root = master
        self.base_frame = tk.Frame(self.root)
        self.base_frame.pack(fill='both', expand=1)
        self.page = self.Page(self.base_frame, self.go_back, self.go_next)
        self.pagei = self.Pagei(self.page.center_frame, self.page.next_button)

    def go_back(self):
        self.pageii.pack_forget()
        self.pagei.pack()
        self.page.back_button.disable()
        self.page.next_button.normal()

    def go_next(self):
        if self.pagei.get_cells() and self.pagei.get_cells() != "选取单元格":
            self.pageii = self.Pageii(self.page.center_frame, self.pagei.address, self.pagei.get_cells())
            self.pageii.change_label(self.pagei.address)
            self.pageii.show_result()

            self.pagei.pack_forget()
            self.page.back_button.normal()
            self.page.next_button.disable()
        else:
            tkinter.messagebox.showinfo("提示", "请按如下格式选取单元格：\n1、A1:A10\n2、A1:C10\n3、A1:C10,E1:E10,F1:F10")

    def destroy(self):
        self.base_frame.destroy()
