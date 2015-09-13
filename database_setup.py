from sqlalchemy import Column, Integer, String, DateTime, Numeric, NVARCHAR, Table, Index, ForeignKey, \
    PrimaryKeyConstraint, ForeignKeyConstraint
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine
from sqlalchemy.orm import relationship, backref

Base = declarative_base()
Engine = create_engine('mssql+pymssql://appadmin:N0v1terp@srvshasql01/R_DHDI_test_0907')


class MapTable(Base):
    __tablename__ = "map_source"
    id = Column(Integer, primary_key=True)
    jde_code = Column(String(100), nullable=False)
    jde_name = Column(String(250), nullable=False)
    kingdee_code = Column(String(100), nullable=False)


class VoucherHead(Base):
    __tablename__ = "c_VoucherHead"
    __table_args__ = (PrimaryKeyConstraint('jde_number', 'period', name='c_VoucherHead_pk'),)
    id = Column(Integer)
    jde_number = Column(NVARCHAR(20), nullable=False)
    line_count = Column(Integer, nullable=False)
    # voucher_number = Column(Integer, nullable=False)
    # voucher_line_number = Column(Integer, nullable=False)
    year = Column(Integer, nullable=False)
    period = Column(Integer, nullable=False)
    voucher_date = Column(DateTime, nullable=False)
    voucher_category = Column(NVARCHAR(20), nullable=False, default='transfer')
    preparer = Column(NVARCHAR(20), nullable=False, default='morningstar')

    def __init__(self, jde_number, line_count, year, period, voucher_date):
        # self.unique_id = unique_id
        self.jde_number = jde_number
        self.line_count = line_count
        self.year = year
        self.period = period
        self.voucher_date = voucher_date
        # self.voucher_category = voucher_category
        # self.preparer = preparer


class VoucherBody(Base):
    __tablename__ = "c_VoucherBody"
    __table_args__ = (ForeignKeyConstraint(['jde_number', 'period'],
                                           ['c_VoucherHead.jde_number', 'c_VoucherHead.period']),)
    id = Column(Integer, primary_key=True)
    jde_number = Column(NVARCHAR(20), nullable=False)
    period = Column(Integer, nullable=False)
    # header_id = Column(Integer, ForeignKey("c_VoucherHead.id"))
    jde_account = Column(NVARCHAR(20), nullable=False)
    currency = Column(NVARCHAR(20), nullable=False)
    exchange_type = Column(NVARCHAR(20), nullable=False, default='公司汇率')
    exchange_rate = Column(Numeric(18, 8), nullable=False)
    amount_for = Column(Numeric(20, 2), nullable=False)
    amount_cny = Column(Numeric(20, 2), nullable=False)
    voucher_description = Column(NVARCHAR(200), nullable=False)
    voucher_heads = relationship("VoucherHead", backref=backref('voucher_bodies', order_by=id))

    def __init__(self, id, jde_number, period, jde_account, currency, exchange_type, exchange_rate, amount_for, amount_cny,
                 voucher_description):
        self.id = id
        self.jde_number = jde_number
        self.period = period
        self.jde_account = jde_account
        self.currency = currency
        self.exchange_type = exchange_type
        self.exchange_rate = exchange_rate
        self.amount_for = amount_for
        self.amount_cny = amount_cny
        self.voucher_description = voucher_description
# c_VoucherHead = Base.metadata.tables['c_VoucherHead']
# Index('ix_JDENumber_Period', Base.metadata.tables['c_VoucherHead'].c.jde_number,
#       Base.metadata.tables['c_VoucherHead'].c.period, unique=True)
# class Voucher(Base):
#     __table__ = Table('t_Voucher', Base.metadata, autoload=True, autoload_with=Engine)
#
#
# class VoucherEntry(Base):
#     __table__ = Table('t_VoucherEntry', Base.metadata, autoload=True, autoload_with=Engine)
class SubSys(Base):
    __table__ = Table('t_SubSys', Base.metadata, autoload=True, autoload_with=Engine)


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


# from sqlalchemy.orm import sessionmaker
# Session=sessionmaker(Engine)
# session=Session()
# a = session.query(VoucherHead)
# print(a[0].voucher_bodies[0].amount_cny)
# for each in a:
#     print(each.voucher_bodies[0].amount_cny)
# drop_table(Engine, 'c_VoucherBody')
# drop_table(Engine, 'c_VoucherHead')
#
# create_table(Engine, 'c_VoucherHead')
# create_table(Engine, 'c_VoucherBody')
# def delete_item(engine, tablename):
#     if if_table_exists(engine, tablename):
#         Base.metadata.tables[tablename].delete()