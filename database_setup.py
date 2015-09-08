from sqlalchemy import Column, Integer, String, DateTime, Numeric, NVARCHAR
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import create_engine


Base = declarative_base()
Engine = create_engine('mssql+pymssql://appadmin:N0v1terp@srvshasql01/DH_DI')


class MapTable(Base):
    __tablename__ = "map_source"
    id = Column(Integer, primary_key=True)
    jde_code = Column(String(100), nullable=False)
    jde_name = Column(String(250), nullable=False)
    kingdee_code = Column(String(100), nullable=False)


class Vouchers(Base):
    __tablename__ = "vouchers"
    inter_id = Column(Integer, primary_key=True)
    # jde_number包括Excel中的Document_Number和流水号；
    jde_number = Column(NVARCHAR(20), unique=True)
    voucher_number = Column(Integer, nullable=False)
    voucher_line_number = Column(Integer, nullable=False)
    voucher_date = Column(DateTime, nullable=False)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)
    voucher_category = Column(NVARCHAR(20), nullable=False)
    jde_account = Column(NVARCHAR(20), nullable=False)
    currency = Column(NVARCHAR(20), nullable=False)
    exchange_type = Column(NVARCHAR(20), nullable=False)
    exchange_rate = Column(Numeric(18, 8), nullable=False)
    amount_for = Column(Numeric(20, 2), nullable=False)
    amount_cny = Column(Numeric(20, 2), nullable=False)
    preparer = Column(NVARCHAR(20), nullable=False)
    voucher_description = Column(NVARCHAR(200), nullable=False)

    def __init__(self, jde_number, voucher_number, voucher_line_number, voucher_date, year, month, voucher_category,
                 jde_account, currency, exchange_type, exchange_rate, amount_for, amount_cny, preparer, voucher_description):
        self.jde_number = jde_number
        self.voucher_number = voucher_number
        self.voucher_line_number = voucher_line_number
        self.voucher_date = voucher_date
        self.year = year
        self.month = month
        self.voucher_category = voucher_category
        self.jde_account = jde_account
        self.currency = currency
        self.exchange_type = exchange_type
        self.exchange_rate = exchange_rate
        self.amount_for = amount_for
        self.amount_cny = amount_cny
        self.preparer = preparer
        self.voucher_description = voucher_description


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


# def delete_item(engine, tablename):
#     if if_table_exists(engine, tablename):
#         Base.metadata.tables[tablename].delete()