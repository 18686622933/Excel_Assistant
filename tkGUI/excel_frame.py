#!/usr/bin/env python3
# -*- coding:utf-8 -*-


import tkinter.messagebox
from tkGUI.base_class import *
from functions.excel_CheckRepeat import *
from functions.excel_Contrast import *
from functions.excel_Merge import *


class Check:
    """查重"""

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
            self.center_frame = father
            self.next_button = next_button

            # 当前页面最底层frame
            self.iframe = tk.Frame(self.center_frame)
            self.iframe.pack(fill="both", expand=1)

            # button_frame、label_frame上下结构
            self.button_frame = tk.Frame(self.iframe)
            self.button_frame.pack(fill="both", expand=1)
            self.button = tk.Button(self.button_frame,
                                    text=" Select the file and make sure the sheet is first. ",
                                    font=('Microsoft YaHei', 26),
                                    cursor="hand2",
                                    relief="flat",
                                    command=self.ask_file)
            self.button.pack(side="bottom")
            self.label_frame = tk.Frame(self.iframe)
            self.label_frame.pack(fill="both", expand=1)
            self.label = tk.Label(self.label_frame, text="", font=('Microsoft YaHei', 15))
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

        def pack_forget(self):
            self.iframe.pack_forget()

        def pack(self):
            self.iframe.pack(fill='both', expand=1)

    class Pageii:
        """第二页完整功能"""

        def __init__(self, father, do_refresh, do_excel):
            self.center_frame = father
            self.iiframe = tk.Frame(self.center_frame)  # 本页面此层frame
            self.iiframe.pack(fill="both", expand=1)
            self.ctrl_frame = tk.Frame(self.iiframe)  # 最上行放置控制按键的frame
            self.info_frame = tk.Frame(self.iiframe)
            self.ctrl_frame.pack(fill="x", side="top")  # 展示信息的活页
            self.info_frame.pack(fill="both", expand=1)

            self.label = tk.Label(self.ctrl_frame, text="Result:")
            self.field = tk.Entry(self.ctrl_frame, )  # 输入字段的Entry
            self.field = TipEntry(self.ctrl_frame, width=26, fg="gray", justify="center")  # 自定义Entry类，包含事件处理
            self.field.set_default_text("筛选查重字段，默认为全部字段")
            self.field.mybind()  # 生成事件
            self.refresh = tk.Button(self.ctrl_frame, text="Refresh", command=do_refresh)  # 刷新按钮
            self.to_excel = tk.Button(self.ctrl_frame, text="to Excel", command=do_excel)  # 输出到Excel按钮

            self.to_excel.pack(side=tk.RIGHT)
            self.refresh.pack(side=tk.RIGHT)
            self.field.pack(side=tk.RIGHT)
            self.label.pack(side=tk.LEFT)

        def create_treeview(self, tv_data):
            self.treeview = TipTreeview(self.info_frame, tv_data)

        def get_field(self):
            return self.field.get()

        def change_label(self, address):
            self.label.configure(text="Result of <%s>'s Checking:" % os.path.basename(address))

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
        try:
            self.data = check_repeat(self.pagei.address)
            self.pageii = self.Pageii(self.page.center_frame, self.do_refresh, self.do_excel)
            self.pageii.change_label(self.pagei.address)

            self.pageii.create_treeview(self.data)
            self.pageii.pack()
            self.pagei.pack_forget()
            self.page.back_button.normal()
            self.page.next_button.disable()
        except:
            tk.messagebox.showerror("提示", "Error")

    def do_refresh(self):
        """点击刷新按钮，则按输入的字段重新筛选数据，先清空treeview再写入新数据"""
        try:
            self.data = check_repeat(self.pagei.address, self.pageii.get_field())
            self.pageii.treeview.clear_out()
            self.pageii.treeview.write(self.data)
        except TypeError:
            pass

    def do_excel(self):
        toexcel = ToExcel(self.pagei.address, "查重结果")
        default_name = toexcel.get_default_name
        save_path = get_save_path(default_name)
        info = toexcel.to_excel(self.data, save_path)
        if info:
            tk.messagebox.showinfo("输出完成", info)

    def destroy(self):
        self.base_frame.destroy()


class Contrast:
    """对比"""

    class Page(Check.Page):
        pass

    class Pagei:
        class PageiHalf:
            def __init__(self, father, button_text, next_button):
                self.button_text = "%s" % button_text
                self.next_button = next_button
                self.root = father
                self.up_frame = tk.Frame(self.root)
                self.middle_frame = tk.Frame(self.root)
                self.down_frame = tk.Frame(self.root)

                self.up_frame.pack(fill="both", expand=1)
                self.middle_frame.pack(fill="x")
                self.down_frame.pack(fill="both", expand=1)

                self.field = TipEntry(self.middle_frame, width=28, fg="gray", justify="center")  # 自定义Entry类，包含事件处理
                self.field.set_default_text("筛选对比字段，默认为全部字段")
                self.field.mybind()  # 生成事件
                self.field.pack()

                self.button = tk.Button(self.up_frame,
                                        text=self.button_text,
                                        width=13,
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
                    self.button.configure(text=self.button_text)
                    self.label.configure(text="")
                    self.next_button.disable()
                else:
                    if base_name and not (base_name.endswith('.xls') or base_name.endswith('.xlsx')):
                        tk.messagebox.showwarning('提示', '请选择正确的Excel文件！')
                    elif len(address) >= 18:
                        show_address = "..." + address[-18:]
                        self.button.configure(text=" Click to re-select ")
                        self.label.configure(text=show_address)
                        self.next_button.normal()

                        self.address = address

            def get_field(self):
                end = self.field.get()
                if end.endswith("默认为全部字段"):
                    return ""
                else:
                    return end

        def __init__(self, father, next_button):
            """还Page的中间容器上添加了上、下两个容易，分别放置选择文件的button，和展示文件路径的label"""
            self.center_frame = father
            self.next_button = next_button
            self.iframe = tk.Frame(self.center_frame, bg='gray')
            self.iframe.pack(fill="both", expand=1)

            # 左右平分frame
            self.l_frame = tk.Frame(self.iframe)
            self.m_frame = tk.Frame(self.iframe)
            self.r_frame = tk.Frame(self.iframe)
            self.l_frame.pack(expand=1, fill="both", side="left", anchor="w")
            self.m_frame.pack(fill="y", side="left", anchor="w")
            self.r_frame.pack(expand=1, fill="both", side="left", anchor="w")
            self.vs = tk.Label(self.m_frame, text="vs", font=('华文新魏', 55))
            self.vs.pack(side=tk.LEFT)

            self.l_half = self.PageiHalf(self.l_frame, "Select one Excel", self.next_button)
            self.r_half = self.PageiHalf(self.r_frame, "Select  Another", self.next_button)

        def pack_forget(self):
            self.iframe.pack_forget()

        def pack(self):
            self.iframe.pack(fill='both', expand=1)

    class Pageii:
        def __init__(self, father, do_excel):
            self.center_frame = father
            self.iiframe = tk.Frame(self.center_frame)  # 本页面此层frame
            self.iiframe.pack(fill="both", expand=1)
            self.ctrl_frame = tk.Frame(self.iiframe)  # 最上行放置控制按键的frame
            self.info_frame = tk.Frame(self.iiframe)
            self.ctrl_frame.pack(fill="x", side="top")  # 展示信息的活页
            self.info_frame.pack(fill="both", expand=1)

            self.label = tk.Label(self.ctrl_frame, text="Result:")
            self.to_excel = tk.Button(self.ctrl_frame, text="to Excel", command=do_excel)  # 输出到Excel按钮

            self.to_excel.pack(side=tk.RIGHT)
            self.label.pack(side=tk.LEFT)

        def create_treeview(self, tv_data):
            self.treeview = TipTreeview(self.info_frame, tv_data)

        def change_label(self, addressi, addressii):
            self.label.configure(text="<%s> vs <%s>:" % (os.path.basename(addressi), os.path.basename(addressii)))

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
        try:
            self.data = contrast(self.pagei.l_half.address, self.pagei.r_half.address, self.pagei.l_half.get_field(),
                                 self.pagei.r_half.get_field())
            self.pageii = self.Pageii(self.page.center_frame, self.do_excel)
            self.pageii.change_label(self.pagei.l_half.address, self.pagei.r_half.address)
            self.pageii.create_treeview(self.data)
            self.pageii.pack()
            self.pagei.pack_forget()

            self.page.back_button.normal()
            self.page.next_button.disable()
        except:
            tkinter.messagebox.showinfo("提示", "所选文件不同，\n则对比两个文件中的首位sheet；\n\n所选文件相同，\n则对比该文件的前两位sheet。")

    def do_excel(self):
        pass
        toexcel = ToExcel(self.pagei.l_half.address, "对比结果")
        default_name = toexcel.get_default_name
        save_path = get_save_path(default_name)
        info = toexcel.to_excel(self.data, save_path)
        if info:
            tk.messagebox.showinfo("输出完成", info)

    def destroy(self):
        self.base_frame.destroy()


class Merge:
    """合并"""

    class Page(Check.Page):
        pass

    class Pagei:
        """第一页完整功能"""

        def __init__(self, father, next_button):
            """还Page的中间容器上添加了上、下两个容易，分别放置选择文件的button，和展示文件路径的label"""
            self.center_frame = father
            self.next_button = next_button

            # 当前页面最底层frame
            self.iframe = tk.Frame(self.center_frame)
            self.iframe.pack(fill="both", expand=1)

            # button_frame、label_frame上下结构
            self.button_frame = tk.Frame(self.iframe)
            self.button_frame.pack(fill="both", expand=1)
            self.button = tk.Button(self.button_frame,
                                    text=" Select the Directory ",
                                    font=('Microsoft YaHei', 26),
                                    width=18,
                                    cursor="hand2",
                                    relief="flat",
                                    command=self.ask_directory)
            self.button.pack(side="bottom")
            self.label_frame = tk.Frame(self.iframe)
            self.label_frame.pack(fill="both", expand=1)
            self.label = tk.Label(self.label_frame, text="", font=('Microsoft YaHei', 15))
            self.label.pack()

            self.directory = ""
            self.excel_list = []

        def ask_directory(self):
            """选择路径函数"""
            directory = filedialog.askdirectory()
            if not directory:
                self.button.configure(text=" Select the Directory ")
                self.label.configure(text="")
                self.next_button.disable()
            else:
                files = [file for file in os.listdir(directory)]
                excel_list = filter(lambda file: file.endswith("xls") or file.endswith("xlsx"), files)
                self.excel_list = sorted(excel_list)

                if self.excel_list:
                    if len(directory) >= 40:
                        short_dir = "..." + directory[-40:]
                        self.label.configure(text="%s" % short_dir, font=('Microsoft YaHei', 13))
                    if len(directory) >= 30:
                        self.label.configure(text="%s" % directory, font=('Microsoft YaHei', 13))
                    else:
                        self.label.configure(text="%s" % directory, font=('Microsoft YaHei', 15))

                    self.next_button.normal()
                    self.button.configure(text=" Click to re-select ")
                    self.directory = directory + "/"
                else:
                    tk.messagebox.showwarning("提示", "请选择Excel文件所在目录")

        def pack_forget(self):
            self.iframe.pack_forget()

        def pack(self):
            self.iframe.pack(fill='both', expand=1)

    class Pageii:
        """第二页完整功能"""

        def __init__(self, father, begin_merging, form_tip):
            self.center_frame = father
            self.iiframe = tk.Frame(self.center_frame)  # 本页面此层frame
            self.iiframe.pack(fill="both", expand=1)
            self.ctrl_frame = tk.Frame(self.iiframe)  # 最上行放置控制按键的frame
            self.info_frame = tk.Frame(self.iiframe)
            self.ctrl_frame.pack(fill="x", side="top")  # 展示信息的活页
            self.info_frame.pack(fill="both", expand=1)

            self.label = tk.Label(self.ctrl_frame)
            self.field = tk.Entry(self.ctrl_frame, )  # 输入字段的Entry
            self.field = TipEntry(self.ctrl_frame, width=18, fg="gray")  # 自定义Entry类，包含事件处理
            self.field.set_default_text("表头所占行数，默认为1")
            self.field.mybind()  # 生成事件

            self.form_frame = tk.LabelFrame(self.ctrl_frame, relief="flat")
            self.checkVar = tk.IntVar()
            self.form = tk.Checkbutton(self.form_frame, text="Form", variable=self.checkVar, onvalue=1, offvalue=0,
                                       command=form_tip)
            self.form.pack()

            self.begin_merging = tk.Button(self.ctrl_frame, text="Begin Merging", command=begin_merging)  # 输出到Excel按钮

            self.begin_merging.pack(side=tk.RIGHT)
            self.form_frame.pack(side=tk.RIGHT)
            # self.form.pack(side=tk.RIGHT)
            self.field.pack(side=tk.RIGHT)
            self.label.pack(side=tk.LEFT)

        def create_listbox(self, lb_data):
            self.listbox = TipListbox(self.info_frame, lb_data)

        def get_field(self):
            if not self.field.get() or self.field.get() == "表头所占行数，默认为1":
                return 1
            else:
                return self.field.get()

        def change_label(self, address):
            self.label.configure(text="%s" % os.path.basename(address))

        def is_form(self):
            return self.checkVar.get()

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
        try:
            self.data = self.pagei.excel_list
            self.pageii = self.Pageii(self.page.center_frame, self.begin_merging, self.form_tip)
            self.pageii.change_label("Select files")

            self.pageii.create_listbox(self.data)
            self.pageii.pack()
            self.pagei.pack_forget()
            self.page.back_button.normal()
            self.page.next_button.disable()

        except:
            tk.messagebox.showerror("提示", "Error")

    def form_tip(self):
        form_tip = "按如下规则制作模板并确保文件名排序在首位\n" \
                   "1、单元格为空：查找其他文件相同位置来填充\n" \
                   "2、字体为红色：仅作为示例展示，数据不保留\n" \
                   "3、背景为黄色：将其他文件相同位置累加填充\n" \
                   "4、背景为绿色：将其他文件相同位置均值填充"

        # 因为checkbutton控件无法在勾选时和取消勾选时执行不同动作，所以要在动作中做判断
        if self.pageii.is_form():
            self.pageii.field.pack_forget()
            tk.messagebox.showinfo("表单合并使用说明", form_tip)
        else:
            self.pageii.field.pack(side=tk.RIGHT)

    def begin_merging(self):

        if self.pageii.is_form():
            save_path = get_save_path("表单合并结果")
            if save_path:
                try:
                    info = table_merge(self.pageii.listbox.get_file_list(self.pagei.directory), save_path)
                    if info:
                        tk.messagebox.showinfo("表单合并完成", info)
                except:
                    tk.messagebox.showerror("提示", "Error")

        else:
            save_path = get_save_path("合并结果")
            if save_path:
                try:
                    info = augment(self.pageii.listbox.get_file_list(self.pagei.directory), save_path,
                                   int(self.pageii.get_field()))
                    if info:
                        tk.messagebox.showinfo("合并完成", info)
                except:
                    tk.messagebox.showerror("提示", "Error")

    def destroy(self):
        self.base_frame.destroy()

# class FormIntegration:
#     """表单整合"""
#
#     def __init__(self, master):
#         self.root = master
#         self.base_frame = tk.Frame(self.root)
#         self.base_frame.pack(fill='both', expand=1)
#
#     def destroy(self):
#         self.base_frame.destroy()
