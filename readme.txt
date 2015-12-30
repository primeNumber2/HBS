将特定的Excel格式的数据，整理后导入数据库；
使用SqlAlchemy、pymssql处理程序和数据库的连接和操作
使用py2exe将程序打包为exe可执行程序；
----------------------------------------------------
main文件是程序的入口；
configuration文件定义了一些常用的字符串，如数据库名称、代码等等
database_setup文件
# 主要创建了中间表VoucherHead和VoucherBody，并且映射了DATABASE中t_Sub_Sys表，用来提取当前期间
# VoucherHead和VoucherBody建立了relation关系，这样就可以通过其中的一张表直接取到另一张表的信息
mysetup文件将python脚本转化为exe文件；
voucher_value文件将特定格式的Excel转化为head 和 body 两类数据；main程序调用这个脚本返回的数据，并插入数据库中；
sample_data是导入excel的模板数据；