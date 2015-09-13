# 这个脚本主要用于针对jde导出的Excel数据进行整理
# 首先提取Excel的第5列数据，通过Document Number来判断哪些行属于同一张凭证；
# 然后剔除只有一行的以及金额为0的数据
import tkinter
from tkinter import filedialog
import xlrd
from datetime import *
import calendar

def choose_file():
    """
    让用户选择要导入的文件，并返回文件的绝对路径
    :return: 导入的Excel文件绝对地址 filename
    """
    root = tkinter.Tk()
    root.withdraw()
    file_name = filedialog.askopenfilename()
    return file_name


def voucher_info():
    """
    将jde中的数据进行转化，用dict的方式存储凭证信息
    整个函数分为两部分，第一部分返回一个hash表，Key是凭证号，value是Excel的行号
    第二部分利用第一部分的结果，将excel的内容存储下来
    :return: 返回一个List,List中每个元素是dictionary,dictionary的key和value是凭证导入的关键字段
    """
    file_name = choose_file()
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

    head = set([])
    body = []
    # voucher_number = 0
    for value in hash_voucher.values():
        # 统计每个凭证有多少行
        # 这个行数在导入t_Voucher时是必须字段，而且可以用来判断哪些凭证只有一行，只有一行的数据是不导入金蝶的；
        line_count = len(value)
        for row_number in value:
            # 每个凭证下，每一行的序号，从1开始排序；最大值和上面的变量line_count是一致的
            line_number = 1
            # JDE导出的数据中，有的数据行金额为0，或者为空,这部分数据需要剔除
            if worksheet.cell_value(row_number, 14) != 0 or worksheet.cell_value(row_number, 14) != '':
                # 取JDE的Excel中的Posting Date，转化为年月日时分秒
                year, month, day, hour, minute, second = xlrd.xldate_as_tuple(worksheet.cell(row_number, 3).value, 0)
                # excel中的batch_number，加上行号是每个月的唯一值，但是跨月时，当存在红冲的业务时，batch_number会重复；
                jde_number = str(worksheet.cell_value(row_number, 6))
                # excel中的batch_number加上行号，加上月份，构成唯一值
                # unique_id = '-'.join([str(jde_number), str(month)])
                voucher_date = datetime(year, month, day)
                # 合并两列形成jde_account
                jde_account = worksheet.cell_value(row_number, 9) + worksheet.cell_value(row_number, 10)
                #  有些分录虽然是外币，但是金额过小，美元栏显示为0；此部分凭证行统一确定为人民币；
                currency = "CNY" if worksheet.cell_value(row_number, 13) == "" or worksheet.cell_value(row_number, 13) == 0\
                    else worksheet.cell_value(row_number, 11)
                # 如果是数值为空，默认返回时string类型的空"",但是需要返回0
                # amount_for指原币金额，amount_cny是人民币金额
                amount_for = worksheet.cell_value(row_number, 13) if worksheet.cell_value(row_number, 13) != "" else 0
                amount_cny = worksheet.cell_value(row_number, 14) if worksheet.cell_value(row_number, 14) != "" else 0
                # 汇率是人民币金额除以原币金额反算出来的，所以和Excel中可能会有尾差
                exchange_rate = 1 if currency == "CNY" else round(amount_cny * 1.0 / amount_for, 8)
                # preparer = "Administrator"
                voucher_description = "-".join(worksheet.row_values(row_number, 5, 9))
                # 用set类型的数据来排除重复值，得到头数据；
                # 此处变量的顺序一定要和database_setup中VoucherHead的类顺序相同
                # 源数据Excel中的日期是单据日期，同一个凭证不同的行日期会不一样，这里统一为该月最后一天
                head.add((jde_number, line_count, year, month, datetime(year, month, calendar.monthrange(year, month))))
                each_line = {'jde_number': jde_number, 'period': month,
                             # 'voucher_date': voucher_date, 'year': year, 'month': month,
                             'line_number': line_number,
                             'jde_account': jde_account, 'currency': currency,
                             'exchange_rate': exchange_rate, 'amount_for': amount_for,
                             'amount_cny': amount_cny, 'voucher_description': voucher_description}
                body.append(each_line)
    return head, body


if __name__ == "__main__":
    vouchers = voucher_info()
    print("总共有凭证",len(vouchers), "行")
    print(vouchers[0])
