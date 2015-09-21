from sqlalchemy import create_engine, MetaData, and_, func, Table
from sqlalchemy.orm import sessionmaker
from voucher_value import voucher_info
from database_setup import VoucherHead, VoucherBody, SubSys, MatchTable, Currency

# NAME对应数据库的名称，CODE是JDE中账套的编码，用来在中间表中识别账套
DI_NAME = "R_DH_DI_0915"
DI_CODE = "01046"
TJ_NAME = "R_DH_TJ_TestPython_0916"
TJ_CODE = "01050"
# 数据导入时，会先将所有的数据导入中间表，然后从中间表中将数据分别导入不同的账套，所以中间表作为数据源是固定的，这里将连接字符串固定下来
Inter_Engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/%s?charset=utf8" % DI_NAME)
Inter_Metadata = MetaData(bind=Inter_Engine)
DB_Session = sessionmaker(bind=Inter_Engine)
Inter_Session = DB_Session()


def get_info(database):
    engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/%s?charset=utf8" % database)
    metadata = MetaData(bind=engine)
    db_session = sessionmaker(bind=engine)
    session = db_session()
    return engine, metadata, session


def insert_interface():
    voucher_head, voucher_body = voucher_info()
    if not voucher_head:
        return 0
    # 查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    engine, metadata, session = get_info(DI_NAME)
    open_year, open_period = session.query(SubSys.Fyear, SubSys.Fperiod).filter(SubSys.Fcheckout == 0).first()
    print("当前会计期间为", open_year, "年", open_period, '月')
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


def interface_to_kingdee(database, account_book):
    """
    取接口表VoucherHead的数据，插入金蝶的凭证表中, 代码分为2大部分，分别对应DH_DI和DH_TJ两个账套
    :return:
    """
    # 查询系统中当前的会计期间，只处理当前会计期间的凭证，之前和之后期间的凭证不处理
    engine, metadata, session = get_info(DI_NAME)
    sub_sys = Table('t_SubSys', metadata, autoload=True, autoload_with=engine)
    open_year, open_period = session.query(sub_sys.c.Fyear, sub_sys.c.Fperiod).filter(sub_sys.c.Fcheckout == 0).first()
    # 清除当前会计期间中凭证头和凭证行的数据,已过账的数据不作处理
    print(database, open_period, open_year)
    conn = engine.connect()
    voucher = Table('t_Voucher', metadata, autoload=True, autoload_with=engine)
    voucher_entry = Table('t_VoucherEntry', metadata, autoload=True, autoload_with=engine)
    delete_id = session.query(voucher.c.FVoucherID).filter(voucher.c.FYear == open_year,
                                                           voucher.c.FPeriod == open_period,
                                                           voucher.c.FPosted == 0)
    stmt = voucher.delete().where(voucher.c.FVoucherID.in_(delete_id))
    conn.execute(stmt)
    stmt = voucher_entry.delete().where(voucher_entry.c.FVoucherID.in_(delete_id))
    conn.execute(stmt)
    print("Voucher delete done")
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
        serial_number = session.query(func.max(voucher.c.FSerialNum)).first()[0]
        # 如果是空账套，没有serial_number，返回的是None，所以要指定初始值为1
        if not serial_number:
            serial_number = 1
        # print(serial_number)
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

        for each_line in voucher_bodies:
            account_id = session.query(MatchTable.F_101).filter(MatchTable.FNumber == each_line.jde_account).first()[0]
            currency_id = session.query(Currency.FCurrencyID).filter(Currency.FNumber == each_line.currency).first()[0]
            fdc = 1 if each_line.amount_cny > 0 else 0
            stmt = voucher_entry.insert().values(FVoucherID=voucher_id, FEntryID=each_line.line_number,
                                                 FExplanation=each_line.voucher_description, FAccountID=account_id,
                                                 FCurrencyID=currency_id, FExchangeRate=each_line.exchange_rate,
                                                 FDC=fdc,
                                                 FAmountFor=abs(each_line.amount_for),
                                                 FAmount=abs(each_line.amount_cny),
                                                 FExchangeRateType=1)
            conn.execute(stmt)



if __name__ == "__main__":
    insert_interface()
    interface_to_kingdee(DI_NAME, DI_CODE)
    print("DI Done")
    interface_to_kingdee(TJ_NAME, TJ_CODE)
    print("TJ Done")
    # interface_to_kingdee(TJ_NAME, TJ_CODE)
    # print("TJ Done")