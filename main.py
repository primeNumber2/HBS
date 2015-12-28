from sqlalchemy import create_engine, MetaData, and_, func, Table
from sqlalchemy.orm import sessionmaker
from voucher_value import voucher_info
from database_setup import Engine, VoucherHead, VoucherBody, SubSys, MatchTable, Currency
from configuation import DB_DI, DB_TJ, DI_CODE, TJ_CODE
from datetime import datetime
from tkinter import messagebox
import pymssql
import _mssql
import logging

logging.basicConfig(level=logging.DEBUG, filename="log.txt")
# 数据导入时，会先将所有的数据导入中间表，然后从中间表中将数据分别导入不同的账套，所以中间表作为数据源是固定的，这里将连接字符串固定下来
DB_Session = sessionmaker(bind=Engine)
Inter_Session = DB_Session()


def insert_interface():
    """
    将Excel的源数据插入中间表，源数据来自于另一个Python脚本voucher_value
    :return:
    """
    voucher_head, voucher_body = voucher_info()
    if not voucher_head:
        return 0
    # 查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    open_year, open_period = Inter_Session.query(SubSys.Fyear, SubSys.Fperiod).filter(SubSys.Fcheckout == 0).first()
    fp = open("log.txt", "a")
    fp.write("\n" + "*" * 40 + "\n")
    fp.write("\n" + str(datetime.now()) + "\n")
    fp.write("开始向接口表插入数据\n")
    print("开始向接口表插入数据\n")
    fp.write("接口会计期间为" + str(open_year) + "年" + str(open_period) + '月\n')
    print("接口会计期间为" + str(open_year) + "年" + str(open_period) + '月\n')
    # 删除当前会计期间中，c_VoucherHead和c_VoucherBody中的所有数据，然后导入新数据
    Inter_Session.query(VoucherBody).filter(
        and_(VoucherBody.year == open_year, VoucherBody.period == open_period)).delete(synchronize_session=False)
    Inter_Session.query(VoucherHead).filter(
        and_(VoucherHead.year == open_year, VoucherHead.period == open_period)).delete(synchronize_session=False)
    Inter_Session.commit()
    fp.write("接口表清除完毕\n")
    for each in voucher_head:
        head_data = VoucherHead(*each)
        Inter_Session.add(head_data)
        Inter_Session.commit()
    for each in voucher_body:
        body_data = VoucherBody(*each)
        Inter_Session.add(body_data)
        Inter_Session.commit()
    fp.write("接口表导入完成, 共导入数据" + str(len(voucher_body)) + "行\n")
    print("接口表导入完成, 共导入数据" + str(len(voucher_body)) + "行\n")
    fp.close()
    return 1


def interface_to_kingdee(database, account_book):
    """
    从接口表中取出数据，插入金蝶的凭证表t_Voucher和t_VoucherEntry
    :param database: 要插入的数据库名称，金蝶中不同的账套对应不同的数据库
    :param account_book: 中间表中账套代码，也是JDE种账套的代码， account_book和上一个参数database是一一对应关系
    :return:
    """
    # 建立engine， metadata 和 session
    engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/%s?charset=utf8" % database)
    metadata = MetaData(bind=engine)
    db_session = sessionmaker(bind=engine)
    session = db_session()
    # 通过金蝶系统的表t_Subsys查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    sub_sys = Table('t_SubSys', metadata, autoload=True, autoload_with=engine)
    open_year, open_period = session.query(sub_sys.c.Fyear, sub_sys.c.Fperiod).filter(sub_sys.c.Fcheckout == 0).first()
    # 清除当前会计期间中凭证头和凭证行的数据,已过账的数据不作处理
    fp = open("log.txt", "a")
    fp.write("-" * 20 + "\n")
    fp.write("开始对账套" + database + "导入凭证\n")
    fp.write("账套" + database + "的当前会计期间为" + str(open_year) + "年" + str(open_period) + "月\n")
    conn = engine.connect()
    # 映射金蝶系统的凭证表t_Voucher和t_VoucherEntry, 后续会在这两张表中做清除和插入
    voucher = Table('t_Voucher', metadata, autoload=True, autoload_with=engine)
    voucher_entry = Table('t_VoucherEntry', metadata, autoload=True, autoload_with=engine)
    # 找到当前会计期间的FVoucherID，然后在凭证表t_Voucher和t_VoucherEntry中删除对应的凭证
    delete_id = session.query(voucher.c.FVoucherID).filter(voucher.c.FYear == open_year,
                                                           voucher.c.FPeriod == open_period,
                                                           voucher.c.FPosted == 0)
    stmt = voucher.delete().where(voucher.c.FVoucherID.in_(delete_id))
    conn.execute(stmt)
    stmt = voucher_entry.delete().where(voucher_entry.c.FVoucherID.in_(delete_id))
    conn.execute(stmt)
    fp.write("凭证删除完毕\n")
    # 从接口表VouchHead中取出头信息，注意：因为针对VoucherHead和VoucherBody建立了Relation关系，
    # 所以可以通过VoucherHead的voucher_bodies属性返回Voucher的Body信息
    # 取出接口表中行合计大于1的凭证
    voucher_head = Inter_Session.query(VoucherHead).filter(
        and_(VoucherHead.year == open_year, VoucherHead.period == open_period, VoucherHead.line_count > 1,
             VoucherHead.account_book == account_book))
    # voucher_number指凭证号，每插入一张凭证，凭证号加1
    voucher_number = 1
    # 遍历VoucherHead复合条件的数据，逐行插入金蝶的凭证头和凭证行
    for each in voucher_head:
        # serial_number是t_Voucher中的一个字段，为递增字段，不受期间影响，每次加一；所以每次插入凭证前，
        # 先取出当前系统的最大值，然后加1插入t_Voucher中
        serial_number = session.query(func.max(voucher.c.FSerialNum)).first()[0]
        # 如果是空账套，没有serial_number，返回的是None，所以要指定初始值为1
        if not serial_number:
            serial_number = 1
        stmt = voucher.insert().values(FVoucherID=-1, FDate=each.voucher_date, FYear=each.year, FPeriod=each.period,
                                       FGroupID=1,
                                       FNumber=voucher_number, FEntryCount=each.line_count,
                                       FDebitTotal=each.total_amount,
                                       FCreditTotal=each.total_amount, FPreparerID=16393, FSerialNum=serial_number + 1,
                                       FTransDate=each.voucher_date, FAttachments=0)
        voucher_number += 1
        conn.execute(stmt)
        # 金蝶凭证头的触发器会在插入后生成凭证头的ID，先取出这个ID，用于凭证行插入
        voucher_id = session.query(voucher.c.FVoucherID).filter(voucher.c.FSerialNum == serial_number + 1).first()[0]
        # 根据Relation关系由VoucherHead直接关联到对应的VoucherBody信息
        voucher_bodies = each.voucher_bodies
        # 针对VoucherHead中的每一行头信息，通过SqlAlchemy的Relation关系，可以直接调用每一行的voucher_bodies属性关联到VoucherBody对象
        for each_line in voucher_bodies:
            # VoucherBody中只有jde_account，此时通过MatchTable转化为account_id
            account_id = session.query(MatchTable.F_101).filter(MatchTable.FNumber == each_line.jde_account).first()[0]
            # 将VoucherBody中的币别代码，如“CNY"和"USD"转化为币别ID，如1， 1001
            currency_id = session.query(Currency.FCurrencyID).filter(Currency.FNumber == each_line.currency).first()[0]
            # fdc是凭证的借贷方，正数在借方，负数在贷方，注意此处规定了借贷方，所以后续金额不再分正负，全部取绝对值
            fdc = 1 if each_line.amount_cny > 0 else 0
            stmt = voucher_entry.insert().values(FVoucherID=voucher_id, FEntryID=each_line.line_number,
                                                 FExplanation=each_line.voucher_description, FAccountID=account_id,
                                                 FCurrencyID=currency_id, FExchangeRate=each_line.exchange_rate,
                                                 FDC=fdc,
                                                 FAmountFor=abs(each_line.amount_for),
                                                 FAmount=abs(each_line.amount_cny),
                                                 FExchangeRateType=1)
            conn.execute(stmt)
    fp.write("账套" + database + "的凭证导入完成," + "共导入凭证" + str(voucher_number) + "张\n")
    fp.close()


if __name__ == "__main__":
    try:
        result = insert_interface()
        if result:
            interface_to_kingdee(DB_DI, DI_CODE)
            interface_to_kingdee(DB_TJ, TJ_CODE)
            messagebox.showinfo("Succeed", "The import is done successfully")
    except:
        messagebox.showerror("Error", "Please see log for error message")
        logging.exception("Error:")