#!/usr/bin/env python3
# -*- coding:utf-8 -*-

import tkinter.messagebox
from functions.excel_base import *
from openpyxl.utils import get_column_letter


class Format:
    def __init__(self, path):
        self.excel = Excel(path)

    def get_cell_list(self, cells):
        pattern1 = r'[，|,]'
        pattern2 = r'[：|:|-]'
        cells = re.split(pattern1, cells)
        cell_list = []
        show_info = "请按如下格式选取单元格：\n1、A1:A10\n2、A1:C10\n3、A1:C10,E1:E10,F1:F10"
        for sub in cells:
            sub_cells = re.split(pattern2, sub)
            try:
                start_row = re.findall(r'[0-9]+', sub_cells[0])[0]
                end_row = re.findall(r'[0-9]+', sub_cells[1])[0]
                start_col = openpyxl.utils.column_index_from_string(re.findall(r'[a-z]+', sub_cells[0], re.I)[0])
                end_col = openpyxl.utils.column_index_from_string(re.findall(r'[a-z]+', sub_cells[1], re.I)[0])
                for col in range(int(start_col), int(end_col) + 1):
                    for row in range(int(start_row), int(end_row) + 1):
                        cell_list.append((col, row))
            except (ValueError, IndexError):
                tkinter.messagebox.showinfo("提示", show_info)

        data_list = [str(self.excel.value(column=i[0], row=i[1])) for i in cell_list]
        data_list = list(filter(lambda x: x != "None", data_list))
        return "('" + "','".join(data_list) + "')"
