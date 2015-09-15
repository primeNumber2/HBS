# 这个脚本主要用于针对jde导出的Excel数据进行整理
# 首先提取Excel的第5列数据，通过Document Number来判断哪些行属于同一张凭证；
# 然后剔除金额为0的数据，将剩下的数据分为头数据和行数据，头数据包含凭证头信息，行数据包含凭证行信息
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


def voucher_info():
    """
    将jde中的数据进行转化，将Excel的内容以list的形式返回
    整个函数分为两部分，第一部分返回一个hash表，Key是凭证号，value是Excel的行号
    第二部分利用第一部分的结果，将excel的内容以两个list的形式进行存储
    :return: 返回两个List,分别存储凭证头和凭证行的信息
    """
    file_name = choose_file()
    if not file_name:
        return 0
    workbook = xlrd.open_workbook(file_name)
    worksheet = workbook.sheet_by_index(0)
    # 取出第6列的数据Document Number，判断哪些行属于同一张凭证
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

    # 新增两个空List, head 和 body，然后将每一行的内容增加到两个list中
    head = []
    body = []
    # 定义凭证号 voucher_number， 每个期间内凭证号是连续的，但是由于只有一行数据的存在，所以这个凭证号不能作为最终系统中的凭证号
    voucher_number = 0
    for value in hash_voucher.values():
        # 每个凭证下，每一行的序号，因为从0开始排序；所以最大值比上面的变量line_count小1
        line_number = 0
        total_amount = 0
        for row_number in value:
            # JDE导出的数据中，有的数据行金额为0，或者为空,这部分数据需要剔除
            if worksheet.cell_value(row_number, 14) != 0 or worksheet.cell_value(row_number, 14) != '':
                # 取JDE的Excel中的Posting Date，转化为年月日时分秒
                year, month, day, hour, minute, second = xlrd.xldate_as_tuple(worksheet.cell(row_number, 3).value, 0)
                # excel中的batch_number，加上行号是每个月的唯一值，但是跨月时，当存在红冲的业务时，batch_number会重复；
                jde_number = str(worksheet.cell_value(row_number, 5))
                # excel中的batch_number加上行号，加上月份，构成唯一值
                # unique_id = '-'.join([str(jde_number), str(month)])
                voucher_date = datetime(year, month, day)
                # 合并两列形成jde_account
                jde_account = worksheet.cell_value(row_number, 9) + worksheet.cell_value(row_number, 10)
                # 有些分录虽然是外币，但是金额过小，美元栏显示为0；此部分凭证行统一确定为人民币；
                currency = "CNY" if worksheet.cell_value(row_number, 13) == "" \
                                    or worksheet.cell_value(row_number, 13) == 0 \
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
                each_line = [jde_number, year, month, line_number, jde_account, currency, exchange_rate, amount_for,
                             amount_cny, voucher_description]
                # each_line = {'jde_number': jde_number, 'period': month,
                # # 'voucher_date': voucher_date, 'year': year, 'month': month,
                #              'line_number': line_number,
                #              'jde_account': jde_account, 'currency': currency,
                #              'exchange_rate': exchange_rate, 'amount_for': amount_for,
                #              'amount_cny': amount_cny, 'voucher_description': voucher_description}
                body.append(each_line)
            line_number += 1
            total_amount += abs(amount_cny)
        voucher_number += 1
        # 取每个凭证的最后一行的部分字段作为头信息
        # 头信息中，最后一列是日期信息，部分当月冲回的凭证，jde_number会不变，这样会造成一个jde_number在一个期间内对应两个日期，
        # 这里默认取最后一行的日期作为头信息
        # head.add((jde_number, voucher_number, line_number, year, month, total_amount/2, voucher_date))
        # dict的Key字段和数据库的字段名保持一致，便于后续代码的赋值
        head.append([jde_number, voucher_number, line_number, year, month, total_amount / 2, voucher_date])
        # head.append(
        # {'jde_number': jde_number, 'voucher_number': voucher_number, 'line_count': line_number, 'year': year,
        #      'period': month, 'total_amount': total_amount / 2, 'voucher_date': voucher_date})
    return head, body


if __name__ == "__main__":
    voucher_head, voucher_body = voucher_info()
    print("总共有凭证头", len(voucher_head), "行")
    print(voucher_head)
