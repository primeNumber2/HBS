# 将generate_voucher中返回的list数据插入数据库
# 尽管数据库中的jde_number字段是unique属性，但是为了避免无意义的报错，
# 在导入前会判断要导入的jde_number是否在数据库中已经存在
# 因为使用py2exe来生成exe可执行文件，所以要import _mssql和pymssql，这是py2exe的一个bug
from sqlalchemy import create_engine, Table, and_, func
from sqlalchemy.orm import sessionmaker

from database_setup import Base, VoucherHead, VoucherBody, SubSys, TVoucher, TVoucherEntry, MapTable, Account, Currency
from voucher_value import voucher_info
from tkinter import messagebox
import _mssql
import pymssql


Engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/R_DH_DI_0915?charset=utf8")
Base.metadata.Bind = Engine
DBSession = sessionmaker(bind=Engine)
session = DBSession()


def insert_interface():
    voucher_head, voucher_body = voucher_info()
    if not voucher_head:
        return 0
    # 查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    open_year, open_period = session.query(SubSys.Fyear, SubSys.Fperiod).filter(SubSys.Fcheckout == 0).first()
    # 删除当前会计期间中，c_VoucherHead和c_VoucherBody中的所有数据，然后导入新数据
    session.query(VoucherBody).filter(and_(VoucherBody.year == open_year, VoucherBody.period == open_period)).delete(
        synchronize_session=False)
    session.query(VoucherHead).filter(and_(VoucherHead.year == open_year, VoucherHead.period == open_period)).delete(
        synchronize_session=False)
    session.commit()
    print("Interface delete done")
    for each in voucher_head:
        head_data = VoucherHead(*each)
        session.add(head_data)
        session.commit()
    print("Heads import are done")
    for each in voucher_body:
        body_data = VoucherBody(*each)
        session.add(body_data)
        session.commit()
    print("Bodies done")
    print("Interface finished")


def interface_to_kingdee():
    """
    取接口表VoucherHead的数据，插入金蝶的凭证表中
    :return:
    """
    # 查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    open_year, open_period = session.query(SubSys.Fyear, SubSys.Fperiod).filter(SubSys.Fcheckout == 0).first()
    # 清除当前会计期间中凭证头和凭证行的数据,已过账的数据不作处理
    delete_id = session.query(TVoucher.FVoucherID).filter(TVoucher.FYear == open_year, TVoucher.FPeriod == open_period,
                                                          TVoucher.FPosted == 0)
    de = delete_id.subquery()
    session.query(TVoucherEntry).filter(TVoucherEntry.FVoucherID.in_(de)).delete(synchronize_session=False)
    session.query(TVoucher).filter(TVoucher.FVoucherID.in_(de)).delete(
        synchronize_session=False)
    print("Voucher delete done")
    # 从接口表VouchHead中取出头信息，注意：因为针对VoucherHead和VoucherBody建立了Relation关系，
    # 所以可以通过VoucherHead的voucher_bodies属性返回Voucher的Body信息
    voucher_head = session.query(VoucherHead).filter(
        and_(VoucherHead.year == open_year, VoucherHead.period == open_period, VoucherHead.line_count > 1))
    voucher_number = 1
    # 遍历VoucherHead复合条件的数据，逐行插入金蝶的凭证头和凭证行
    for each in voucher_head:
        serial_number = session.query(func.max(TVoucher.FSerialNum)).first()[0]
        # 如果是空账套，没有serial_number，返回的是None，所以要指定初始值为1
        if not serial_number:
            serial_number = 1
        head_data = TVoucher(FVoucherID=-1, FDate=each.voucher_date, FYear=each.year, FPeriod=each.period, FGroupID=1,
                             FNumber=voucher_number, FEntryCount=each.line_count, FDebitTotal=each.total_amount,
                             FCreditTotal=each.total_amount, FPreparerID=16393, FSerialNum=serial_number + 1,
                             FTransDate=each.voucher_date, FAttachments=0)
        voucher_number += 1
        session.add(head_data)
        session.flush()
        # 金蝶凭证头的触发器会在插入后生成凭证头的ID，先取出这个ID，用于凭证行插入
        voucher_id = session.query(TVoucher.FVoucherID).filter(TVoucher.FSerialNum == serial_number + 1).first()[0]
        # 根据Relation关系由VoucherHead直接关联到对应的VoucherBody信息
        voucher_bodies = each.voucher_bodies

        for each_line in voucher_bodies:
            account_code = \
                session.query(MapTable.kingdee_code).filter(MapTable.jde_code == each_line.jde_account).first()[0]
            account_id = session.query(Account.FAccountID).filter(Account.FNumber == account_code).first()[0]
            currency_id = session.query(Currency.FCurrencyID).filter(Currency.FNumber == each_line.currency).first()[0]
            fdc = 1 if each_line.amount_cny > 0 else 0
            body_data = TVoucherEntry(FVoucherID=voucher_id, FEntryID=each_line.line_number,
                                      FExplanation=each_line.voucher_description, FAccountID=account_id,
                                      FCurrencyID=currency_id, FExchangeRate=each_line.exchange_rate, FDC=fdc,
                                      FAmountFor=abs(each_line.amount_for), FAmount=abs(each_line.amount_cny),
                                      FExchangeRateType=1)
            session.add(body_data)
            session.flush()
        session.commit()


if __name__ == "__main__":
    insert_interface()
    interface_to_kingdee()
    messagebox.showinfo("Import", "Done")