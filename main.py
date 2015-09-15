# 将generate_voucher中返回的list数据插入数据库
# 尽管数据库中的jde_number字段是unique属性，但是为了避免无意义的报错，
# 在导入前会判断要导入的jde_number是否在数据库中已经存在
# 因为使用py2exe来生成exe可执行文件，所以要import _mssql和pymssql，这是py2exe的一个bug
from sqlalchemy import create_engine, Table, and_, func
from sqlalchemy.orm import sessionmaker

from database_setup import Base, VoucherHead, VoucherBody, SubSys, TVoucher, TVoucherEntry
from voucher_value import voucher_info
from tkinter import messagebox
import _mssql
import pymssql


Engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/R_DHDI_test_0907?charset=utf8")
Base.metadata.Bind = Engine
DBSession = sessionmaker(bind=Engine)
session = DBSession()


def insert_interface():
    voucher_head, voucher_body = voucher_info()
    if not voucher_head:
        return 0
    # 查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    open_year, open_period = session.query(SubSys.Fyear, SubSys.Fperiod).filter(SubSys.Fcheckout == 0).first()
    # print(open_year, open_period)
    # 删除当前会计期间中，c_VoucherHead和c_VoucherBody中的所有数据，然后导入新数据
    session.query(VoucherBody).filter(and_(VoucherBody.year == open_year, VoucherBody.period == open_period)).delete(
        synchronize_session=False)
    session.query(VoucherHead).filter(and_(VoucherHead.year == open_year, VoucherHead.period == open_period)).delete(
        synchronize_session=False)
    # session.query(TVoucher).filter(and_(TVoucher.FYear == open_year, TVoucher.FPeriod == open_period)).delete(
    # synchronize_session=False)
    session.commit()
    print("Delete done")
    for each in voucher_head:
        head_data = VoucherHead(*each)
        # head_data = VoucherHead(jde_number=each['jde_number'], voucher_number=each['voucher_number'],
        # line_count=each['line_count'], year=each['year'], period=each['period'],
        # total_amount=each['total_amount'], voucher_date=each['voucher_date'])
        # serial_number = session.query(func.max(TVoucher.FSerialNum)).first()[0] + 1
        # kingdee_data = TVoucher(FVoucherID=-1, FDate=each['voucher_date'], FYear=each['year'], FPeriod=each['period'],
        # FGroupID=1, FNumber=each['voucher_number'], FAttachments=1,
        #                         FEntryCount=each['line_count'], FDebitTotal=each['total_amount'],
        #                         FCreditTotal=each['total_amount'], FPreparerID=16393, FSerialNum=serial_number,
        #                         FTransDate=each['voucher_date'])
        session.add(head_data)
        # session.flush()
        # voucher_id = session.query(TVoucher.FVoucherID).filter(
        #     and_(TVoucher.FYear == each['year'], TVoucher.FPeriod == each['period'],
        #          TVoucher.FNumber == each['voucher_number'])).first()[0]
        # print(voucher_id)

        session.commit()

    print("Heads import are done")
    for each in voucher_body:
        body_data = VoucherBody(*each)
        session.add(body_data)
        session.commit()
    print("Bodies done")


def interface_to_kingdee():
    """
    取接口表VoucherHead的数据，插入金蝶的凭证表中
    :return:
    """
    # 查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    open_year, open_period = session.query(SubSys.Fyear, SubSys.Fperiod).filter(SubSys.Fcheckout == 0).first()
    # 清除当前会计期间中凭证头和凭证行的数据
    session.query(TVoucher).filter(TVoucher.FYear == open_year, TVoucher.FPeriod == open_period).delete(synchronize_session=False)
    voucher_head = session.query(VoucherHead).filter(
        and_(VoucherHead.year == open_year, VoucherHead.period == open_period, VoucherHead.line_count > 1))
    voucher_number = 1

    # print(voucher_head[0].year)
    # print(TVoucher.__table__.columns)
    print(TVoucherEntry.__table__.columns)
    for each in voucher_head:
        serial_number = session.query(func.max(TVoucher.FSerialNum)).first()[0]
        head_data = TVoucher(FVoucherID=-1, FDate=each.voucher_date, FYear=each.year, FPeriod=each.period, FGroupID=1,
                             FNumber=voucher_number, FEntryCount=each.line_count, FDebitTotal=each.total_amount,
                             FCreditTotal=each.total_amount, FPreparerID=16393, FSerialNum=serial_number + 1, FTransDate=each.voucher_date)
        voucher_number += 1
        session.add(head_data)
        session.flush()
        voucher_id = session.query(TVoucher.FVoucherID).filter(TVoucher.FSerialNum == serial_number + 1).first()[0]

        voucher_bodies = each.voucher_bodies
        for each_line in voucher_bodies:
            # body_data = TVoucherEntry(FVoucherID=voucher_id, FEntryID=each_line.line_number)
            print(each_line.line_number, each_line.amount_cny)

        session.commit()
        print(voucher_id)
        # voucher_id = serial_number.query(TVoucher.FVoucherID).filter(TVoucher.serial_number == serial_number + 1)

if __name__ == "__main__":
    # insert_interface()
    interface_to_kingdee()
    # serial_number = session.query(func.max(TVoucher.FSerialNum))
    # print(serial_number.first()[0])
    # print(TVoucher.__table__.columns)
    # t_identity = Table('t_Identity', Base.metadata, autoload=True, autoload_with=Engine)
    # print(t_identity.columns)
    # next_value = session.query(t_identity.c.FNext).filter(t_identity.c.FName == 't_Voucher').first()[0]
    # next_value += 1
    # data = TVoucher(FVoucherID=-1, FDate='2015-09-07', FYear=2015, FPeriod=9, FGroupID=1, FNumber=11, FAttachments=1,
    # FEntryCount=10, FDebitTotal=100, FCreditTotal=100, FPreparerID=16393, FSerialNum=1, FTranType=0,
    # FTransDate='2015-09-07')
    # session.add(data)
    # session.flush()
    # session.commit()
    # voucher_id = session.query(TVoucher.FVoucherID).filter(and_(TVoucher.FYear == 2015, TVoucher.FPeriod == 9,
    # TVoucher.FGroupID == 1,
    # TVoucher.FNumber == 11)).first()[0]
    # print(voucher_id)
    # session.commit()
# 0,default,'2015-09-07',2015,9,
# 1,1,null,null,1,
# 20,4260719.18,4260719.18,null,0,
# 0,16393,-1,-1,-1,
# null,1,null,null,1781,
# 0,'2015-09-07',-1,-1,'',
# default,default
# insert_date()
# voucher_head = session.query(VoucherHead).filter(VoucherHead.jde_number == '74997')
# for c in voucher_head:
# print(c.columns)
# print(SubSys.__table__.columns)
# session.query(SubSys)
# for each in session.query(SubSys):
# print(each.Fsubsysid, each.Fnumber, each.Fname, each.Fused, each.Fperiodsynch, each.Fyear, each.Fyear, each.Fperiod, each.Fcheckout)
# data = SubSys(Fsubsysid=1, Fnumber='Gl', Fname='会计总账管理系统', Fused=1, Fperiodsynch=1, Fyear=2015, Fperiod=9, Fcheckout=0)
# session.add(data)
# session.flush()

# data = SubSys(Fnumber=)
# data = SubSys()
# print(voucher_head)
# for each in voucher_head:
# voucher_bodies = each.voucher_bodiesb
# for line in voucher_bodies:
# print(line.amount_cny)
# # print(each.voucher_bodies.amount_cny)
# voucher_body = voucher_info()[1]
# for each in voucher_head:
# data = VoucherHead()


# print(SubSys.__table__.columns)
# open_period = session.query(SubSys.Fperiod).filter(SubSys.Fcheckout == 0)
#
# print(open_period.first()[0])
# values = session.query(SubSys)
# for each in values:
# print(each.Fperiod, each.Fcheckout)
# vouchers = voucher_info()
# open_period = session.query(SubSys.FPeriod).filter_by()
# print(SubSys.__table__.columns)
# for each in open_period:
# print(each.Fperiod)
# print(open_period)
# for each in open_period:
# print(each.FPeriod)
# for each in open_period:
# print(each)
# values = session.query(VoucherHeader).all()
# for each_value in values:

# jde_numbers = [voucher.jde_number for voucher in session.query(Vouchers).all()]
# num = insert_data(vouchers_data, jde_numbers)
# messagebox.showinfo("succeed", str(num) + ' rows were imported' )
