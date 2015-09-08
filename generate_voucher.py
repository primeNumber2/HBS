# 这个脚本主要用于针对jde导出的Excel数据进行整理
# 首先提取Excel的第5列数据，通过Document Number来判断哪些行属于同一张凭证；
# 然后剔除只有一行的以及金额为0的数据
import tkinter
from tkinter import filedialog
import xlrd
from datetime import *


def choose_file():
    """
    让用户选择要导入的文件，并返回文件的绝对路径
    :return: 导入的Excel文件绝对地址 filename
    """
    root = tkinter.Tk()
    root.withdraw()
    file_name = filedialog.askopenfilename()
    return file_name


def generate_voucher(file_name):
    """
    将jde中的数据进行转化，用dict的方式存储凭证号和内容
    整个函数分为两部分，第一部分返回一个hash表，Key是凭证号，value是Excel的行号
    第二部分返回一个List，将凭证号和Excel行中内容合并到一起
    :param filename: 要处理的Excel对象
    :return: 一个List,返回可以导入的凭证信息
    """
    workbook = xlrd.open_workbook(file_name)
    worksheet = workbook.sheet_by_index(0)
    # 取出第5列的数据Document Number，判断哪些行属于同一张凭证
    cells = worksheet.col_slice(5, 1)
    hash_voucher = {}
    row_number = 1
    for cell in cells:
        # 如果hash的Key表中已经存在该Document Number，将行号row number加入hash表的value中，
        # 如果没有存在该Document Number，将这个Document Number新增为Key
        if cell.value in hash_voucher:
            hash_voucher[cell.value].append(row_number)
        else:
            hash_voucher[cell.value] = [row_number]
        row_number += 1
    # 针对hash_voucher进行筛选，只有凭证行大于2行的凭证，才会被导入
    hash_voucher_validated = {}
    for key in hash_voucher.keys():
        if len(hash_voucher[key]) > 1:
            hash_voucher_validated[key] = hash_voucher[key]



    vouchers = []
    voucher_number = 1

    for rows in hash_voucher_validated.values():
        # 每导入一行数据，凭证行号需要加1；每张凭证做完的时候，凭证号归0
        voucher_line_number = 0

        for row_number in rows:
            # JDE导出的数据中，大量的数据行金额为0，要把这部分数据剔除
            if abs(worksheet.cell(row_number, 14).value) > 0:
                # jde中的Document_Number相当于凭证号，唯一值，不可重复导入；
                # 此处是变通处理，用Document_number加凭证行号作为唯一值；
                jde_number = str(worksheet.cell_value(row_number, 5)) + "-" + str(voucher_line_number)
                # 取JDE的Excel中的Posting Date，转化为年月日时分秒
                year, month, day, hour, minute, second = xlrd.xldate_as_tuple(worksheet.cell(row_number, 3).value, 0)
                voucher_date = datetime(year, month, day)
                voucher_category = "Transfer"
                jde_account = worksheet.cell_value(row_number, 9) + worksheet.cell_value(row_number, 10)
                #  有些分录虽然是外币，但是金额过小，美元栏显示为0；此部分凭证行统一确定为人民币；
                currency = "CNY" if worksheet.cell_value(row_number, 13) == "" or worksheet.cell_value(row_number, 13) == 0\
                    else worksheet.cell_value(row_number, 11)
                # 如果是数值为空，默认返回时string类型的空"",但是需要返回0
                amount_for = worksheet.cell_value(row_number, 13) if worksheet.cell_value(row_number, 13) != "" else 0
                amount_cny = worksheet.cell_value(row_number, 14) if worksheet.cell_value(row_number, 14) != "" else 0
                exchange_type = "公司汇率"
                exchange_rate = 1 if currency == "CNY" else round(amount_cny * 1.0 / amount_for, 8)
                preparer = "Administrator"
                voucher_description = "-".join(worksheet.row_values(row_number, 5, 9))

                voucher_line = [jde_number, voucher_number, voucher_line_number, voucher_date, year, month, voucher_category,
                                jde_account, currency, exchange_type, exchange_rate, amount_for, amount_cny, preparer,
                                voucher_description]
                vouchers.append(voucher_line)
                # 凭证行号加1
                voucher_line_number += 1
        # 凭证号加1
        voucher_number += 1
    return vouchers


def get_vouchers():
    filename = choose_file()
    if filename:
        vouchers = generate_voucher(filename)
        return vouchers

if __name__ == "__main__":
    vouchers = get_vouchers()
    print("总共有凭证" ,len(vouchers), "行")
    print(vouchers[0])
    # for
    # for num in range(10000):
    #     print(vouchers[num])