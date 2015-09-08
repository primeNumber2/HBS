# 将generate_voucher中返回的list数据插入数据库
# 尽管数据库中的jde_number字段是unique属性，但是为了避免无意义的报错，
#   在导入前会判断要导入的jde_number是否在数据库中已经存在
# 因为使用py2exe来生成exe可执行文件，所以
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

from database_setup import Base, Vouchers
from generate_voucher import get_vouchers
from tkinter import messagebox
import _mssql
import pymssql


Engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/DH_DI?charset=utf8")
Base.metadata.Bind = Engine
DBSession = sessionmaker(bind=Engine)
session = DBSession()


def insert_data(vouchers_data, jde_numbers):
    num = 0
    for vouchers in vouchers_data:
        if vouchers[0] not in jde_numbers:
            data = Vouchers(*vouchers)
            session.add(data)
            num += 1
    session.commit()
    session.close()
    return num

if __name__ == "__main__":
    vouchers_data = get_vouchers()
    jde_numbers = [voucher.jde_number for voucher in session.query(Vouchers).all()]
    num = insert_data(vouchers_data, jde_numbers)
    messagebox.showinfo("succeed", str(num) + ' rows were imported' )
