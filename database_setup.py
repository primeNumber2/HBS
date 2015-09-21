# 这张表主要创建了中间表VoucherHead和VoucherBody，并且映射了DATABASE中t_Sub_Sys表，用来提取当前期间
# VoucherHead和VoucherBody建立了relation关系，这样就可以通过其中的一张表直接取到另一张表的信息
from sqlalchemy import Column, Integer, DateTime, Numeric, NVARCHAR, Table, \
    PrimaryKeyConstraint, ForeignKeyConstraint
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine
from sqlalchemy.orm import relationship, backref
from datetime import datetime

Base = declarative_base()
Engine = create_engine('mssql+pymssql://appadmin:N0v1terp@srvshasql01/R_DH_DI_0915')
DATABASE = 'R_DH_DI_0915'


class VoucherHead(Base):
    __tablename__ = "c_VoucherHead"
    # 导入数据没有唯一的表示标识ID，唯一可以确定的是，在同一个期间内，jde_number不会重复出现
    __table_args__ = (PrimaryKeyConstraint('account_book', 'jde_number', 'year', 'period', name='c_VoucherHead_pk'),
    )
    account_book = Column(NVARCHAR(20), nullable=False)
    jde_number = Column(NVARCHAR(20), nullable=False)
    # voucher_number = Column(Integer, nullable=False)
    line_count = Column(Integer, nullable=False)
    year = Column(Integer, nullable=False)
    period = Column(Integer, nullable=False)
    total_amount = Column(Numeric(20, 2), nullable=False)
    voucher_date = Column(DateTime, nullable=False)
    voucher_category = Column(NVARCHAR(20), nullable=False, default='transfer')
    preparer = Column(NVARCHAR(20), nullable=False, default='morningstar')
    create_date = Column(DateTime, default=datetime.now())

    def __init__(self, account_book, jde_number, line_count, year, period, total_amount, voucher_date):
        self.account_book = account_book
        self.jde_number = jde_number
        self.line_count = line_count
        self.year = year
        self.period = period
        self.total_amount = total_amount
        self.voucher_date = voucher_date


class VoucherBody(Base):
    __tablename__ = "c_VoucherBody"
    __table_args__ = (ForeignKeyConstraint(['account_book', 'jde_number', 'year', 'period'],
                                           ['c_VoucherHead.account_book', 'c_VoucherHead.jde_number',
                                            'c_VoucherHead.year', 'c_VoucherHead.period']),
                      PrimaryKeyConstraint('account_book', 'jde_number', 'year', 'period', 'line_number',
                                           name='c_VoucherBody_pk'))
    # id = Column(Integer, primary_key=True)
    account_book = Column(NVARCHAR(20), nullable=False)
    jde_number = Column(NVARCHAR(20), nullable=False)
    year = Column(Integer, nullable=False)
    period = Column(Integer, nullable=False)
    line_number = Column(Integer, nullable=False)
    # header_id = Column(Integer, ForeignKey("c_VoucherHead.id"))
    jde_account = Column(NVARCHAR(20), nullable=False)
    currency = Column(NVARCHAR(20), nullable=False)
    exchange_type = Column(NVARCHAR(20), nullable=False, default='公司汇率')
    exchange_rate = Column(Numeric(18, 8), nullable=False)
    amount_for = Column(Numeric(20, 2), nullable=False)
    amount_cny = Column(Numeric(20, 2), nullable=False)
    voucher_description = Column(NVARCHAR(200), nullable=False)
    voucher_heads = relationship("VoucherHead", backref=backref('voucher_bodies'))
    create_date = Column(DateTime, default=datetime.now())

    def __init__(self, account_book, jde_number, year, period, line_number, jde_account, currency,
                 exchange_rate, amount_for, amount_cny, voucher_description):
        self.account_book = account_book
        self.jde_number = jde_number
        self.year = year
        self.period = period
        self.line_number = line_number
        self.jde_account = jde_account
        self.currency = currency
        self.exchange_rate = exchange_rate
        self.amount_for = amount_for
        self.amount_cny = amount_cny
        self.voucher_description = voucher_description


class SubSys(Base):
    __table__ = Table('t_SubSys', Base.metadata, autoload=True, autoload_with=Engine)


# class TVoucher(Base):
#     # 对应金蝶系统的凭证头表
#     __table__ = Table('t_Voucher', Base.metadata,
#                       autoload=True, autoload_with=Engine, implicit_returning=False)
#
#
# class TVoucherEntry(Base):
#     # 对应金蝶系统的凭证行表
#     __table__ = Table('t_VoucherEntry', Base.metadata,
#                       autoload=True,
#                       autoload_with=Engine)


class MatchTable(Base):
    # 金蝶系统内的科目映射表
    __table__ = Table('t_Item_3001', Base.metadata, autoload=True, autoload_with=Engine)


class Account(Base):
    __table__ = Table('t_Account', Base.metadata, autoload=True, autoload_with=Engine)


class Currency(Base):
    __table__ = Table('t_Currency', Base.metadata, autoload=True, autoload_with=Engine)


# class MapTable(Base):
#     __table__ = Table('map_source', Base.metadata, autoload=True, autoload_with=Engine)


def if_table_exists(engine, tablename):
    return engine.dialect.has_table(engine.connect(), tablename)


def drop_table(engine, tablename):
    if if_table_exists(engine, tablename):
        Base.metadata.tables[tablename].drop(engine)
        print('Table', tablename.upper(), "is dropped")
    else:
        pass


def create_table(engine, tablename):
    if not if_table_exists(engine, tablename):
        Base.metadata.tables[tablename].create(engine)
        print('Table', tablename.upper(), "is created")
    else:
        print("Already exists")


        # drop_table(Engine, 'c_VoucherBody')
        # drop_table(Engine, 'c_VoucherHead')
        # create_table(Engine, 'c_VoucherHead')
        # create_table(Engine, 'c_VoucherBody')
        # # def delete_item(engine, tablename):
        #     if if_table_exists(engine, tablename):
        #         Base.metadata.tables[tablename].delete()