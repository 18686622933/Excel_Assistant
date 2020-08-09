#!/usr/bin/env python3
# -*- coding:utf-8 -*-


from functions.excel_base import *


def contrast(path_i, path_ii, column_i="", column_ii=""):
    # 将两个表中的数据放入set中，以加快比对速度
    def get_data_set(excel, column_list):
        """
        将一个sheet中的数据放到由元组(每行数据为一个元组)组成的集合中
        :param excel: Excel类
        :param column_list: 要取的字段列表
        :return: {{字段A,字段B,...},{字段A,字段B,...},{字段A,字段B,...},...}
        """
        sheet_data = set()
        for r in range(1, excel.max_row + 1):
            r_data = tuple([excel.value(column, r) for column in column_list])  # 按提供字段获取一行数据，再转为tuple格式
            sheet_data.add(r_data)  # 只有tuple格式才可以添加到set中
        return sheet_data

    error_text = "两个表所选字段数量需要一致！"
    if len(column_i) != len(column_ii):  # 如果两个表提供的字段数量不一致，则返回错误提示
        print(error_text)
        return error_text
    else:  # 将要对比的两个表构造为两个Excel类实例
        excel_i = Excel(path_i)
        if path_i == path_ii:  # 如果所选为同一文件，则对比该文件的前两个sheet
            excel_ii = Excel(path_ii, 2)
        else:  # 否则对比两个表的第一个sheet
            excel_ii = Excel(path_ii)

    # 通过get_column_list()函数将用户填入的数字或字母字段转换为数字索引
    def all_columns(excel):
        column_list = [str(column) for column in range(1, excel.max_column + 1)]
        return ",".join(column_list)

    if column_i == "":
        column_i = all_columns(excel_i)
    if column_ii == "":
        column_ii = all_columns(excel_ii)

    column_num_list_i = get_column_list(column_i)
    column_num_list_ii = get_column_list(column_ii)
    column_letter_list_i = get_column_list(column_i, letter_or_num="a")
    column_letter_list_ii = get_column_list(column_ii, letter_or_num="a")
    column_dict = dict(zip(column_letter_list_i, column_letter_list_ii))

    # 创建序号参数和构造表头
    num = 1
    contrast_result = [["序号", "问题描述", "原表中行号"] + ["表1:%s列 表2:%s列" % (k, v) for k, v in column_dict.items()]]

    # 先创建表2的set对象，然后遍历表1，生成表1set对象的同时判断每行数据是否存在于表2中，不存在则放到contrast_result列表中。
    sheet_data_ii = get_data_set(excel_ii, column_num_list_ii)
    sheet_data_i = set()
    for row in range(1, excel_i.max_row + 1):
        row_data = tuple([excel_i.value(column, row) for column in column_num_list_i])
        sheet_data_i.add(row_data)
        if row_data not in sheet_data_ii:
            contrast_result.append([num] + ["在表1不在表2"] + [row] + list(row_data))
            num += 1
    # 上面循环完成了创建表1的set和将表1与表2set进行对比，下面遍历表2与表1的set进行对比。
    for row in range(1, excel_ii.max_row + 1):
        row_data = tuple([excel_ii.value(col, row) for col in column_num_list_ii])
        if row_data not in sheet_data_i:
            contrast_result.append([num] + ["在表2不在表1"] + [row] + list(row_data))
            num += 1

    return contrast_result


if __name__ == '__main__':
    file_path1 = "//test3.xlsx"
    file_path2 = "//test3.xlsx"
    different_data = contrast(file_path1, file_path2)
    print(different_data)
