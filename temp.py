# 将generate_voucher中返回的list数据插入数据库
# 尽管数据库中的jde_number字段是unique属性，但是为了避免无意义的报错，
# 在导入前会判断要导入的jde_number是否在数据库中已经存在
# 因为使用py2exe来生成exe可执行文件，所以要import _mssql和pymssql，这是py2exe的一个bug
from sqlalchemy import create_engine, MetaData, Table, and_
from sqlalchemy.orm import sessionmaker

Engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/R_DH_DI_0915?charset=utf8")
metadata = MetaData(bind=Engine)
DBSession = sessionmaker(bind=Engine)
session = DBSession()
conn = Engine.connect()
map_table = Table('map_source', metadata, autoload=True, autoload_with=Engine)
t_item_3001 = Table('t_Item_3001', metadata, autoload=True, autoload_with=Engine)
t_item = Table('t_Item', metadata, autoload=True, autoload_with=Engine)
t_account = Table('t_Account', metadata, autoload=True, autoload_with=Engine)
print(map_table.columns)
print(t_item_3001.columns)
# values = session.query(map_table.c.jde_code, map_table.c.jde_name)
# for each in values:
# print(each)
values = session.query(t_item.c.FItemID, t_item.c.FNumber).filter(t_item.c.FItemClassID == 3001)
# for each in values:
# print(each)
print(values[0])
result = session.query(map_table.c.jde_code, map_table.c.AccountID, t_item.c.FItemID, map_table.c.jde_name, t_account.c.FName).join(
    t_item, t_account, and_(map_table.c.jde_code == t_item.c.FNumber, map_table.c.AccountID == t_account.c.FAccountID, t_item.c.FItemClassID == 3001,  ))
print(result[0])
# for each in result:
# ins = t_item_3001.insert().values(FItemID=each[2], F_101=each[1], FNumber=each[0], FName=each[3])
#     conn.execute(ins)