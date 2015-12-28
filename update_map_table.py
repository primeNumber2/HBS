from sqlalchemy import create_engine, MetaData, Table
from sqlalchemy.orm import sessionmaker

Engine = create_engine("mssql+pymssql://appadmin:N0v1terp@srvshasql01/DH_DI?charset=utf8")
metadata = MetaData(bind=Engine)
DBSession = sessionmaker(bind=Engine)
session = DBSession()
conn = Engine.connect()

map_table = Table('map_source', metadata, autoload=True, autoload_with=Engine)
t_item_3001 = Table('t_Item_3001', metadata, autoload=True, autoload_with=Engine)
t_item = Table('t_Item', metadata, autoload=True, autoload_with=Engine)

print(map_table)
print(t_item_3001.columns)
print(t_item.columns)

source = session.query(map_table.c.jde_code, map_table.c.jde_name, map_table.c.kingdee_code, map_table.c.AccountID)
# item_id = 1
for each in source:
    ins = t_item.insert().values(FItemID=-1, FItemClassID=3001, FNumber=each[0], FParentID=0, FLevel=1, FDetail=1,
                                 FName=each[1], FFullNumber=each[0], FShortNumber=each[0], FFullName=each[1])
    conn.execute(ins)
# #     item_id += 1
# constraint_type
# DEFAULT on column FAccessory
# DEFAULT on column FBrNo
# DEFAULT on column FDeleted
# DEFAULT on column FDiff
# DEFAULT on column FExternID
# DEFAULT on column FFullNumber
# DEFAULT on column FGRCommonID
# DEFAULT on column FGrControl
# DEFAULT on column FHavePicture
# DEFAULT on column FItemID
# DEFAULT on column FSystemType
# DEFAULT on column FUnUsed
# DEFAULT on column FUseSign
# DEFAULT on column FLevel
# DEFAULT on column UUID
