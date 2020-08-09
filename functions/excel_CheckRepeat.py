#!/usr/bin/env python3
# -*- coding:utf-8 -*-

from functions.excel_base import *


def check_repeat(path, columns=""):
    """
    对指定excel文件的sheet1进行查重，函数返回查重结果，可以指定字段，默认为全部字段
    :param path: 指定excel
    :param columns: 指定字段，默认为全部字段
    :return: 返回查重结果，可以使用to_excel()函数写入新的excel中
    """
    excel = Excel(path)

    sheet_data = {}
    check_result = [["序号", "行号"] + list(excel.row_data_tuple(1, columns))]

    num = 1
    for row in range(1, excel.max_row + 1):
        row_data = excel.row_data_tuple(row, columns)
        if row_data not in sheet_data:
            sheet_data[row_data] = row
        else:
            check_result.append([num] + [sheet_data[row_data]] + list(row_data))
            num += 1
            check_result.append([num] + [row] + list(row_data))
            num += 1
    return check_result


if __name__ == '__main__':
    file_path1 = "//test1.xls"
    file_path2 = "//test3.xlsx"
    repeat_data = check_repeat(file_path2, "1567")
    # print(check_repeat(file_path2))
    for i in check_repeat(file_path2):
        print(i)

    # to_excel(repeat_data, file_path2, "查重结果")
