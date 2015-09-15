from datetime import *
import calendar
a = [1,2,3]
b = [ 2,7, 9]
for i , j  in zip(a, b):
    print(i,j)

# a = set([1,2, 2])
# b = set([[1,2], [2,4] ])
# a.add(3)
# print(a)
# b.add((2,4))
# print(b)
# for i in b:
#     print(i)
# print('.'.join(['a', 'b', 'c']))

# from sqlalchemy import create_engine, Column, Integer, String, ForeignKey
# from sqlalchemy.ext.declarative import declarative_base
# from sqlalchemy.orm import sessionmaker, relationship, backref
# from sqlalchemy.orm.exc import MultipleResultsFound, NoResultFound
# import database_setup as ds
#
# Engine = create_engine('mssql+pymssql://appadmin:N0v1terp@srvshasql02/Test')
# Base = declarative_base(Engine)
# Session = sessionmaker(Engine)
# session = Session()
#
#
# class Person(Base):
#     __tablename__ = "persons"
#     id = Column(Integer, primary_key=True)
#     name = Column(String(100))
#     age = Column(Integer)
#
#     def __repr__(self):
#         return "id = %s, name = %s, age = %s" % (self.id, self.name, self.age)
#
#
# class Department(Base):
#     __tablename__ = "departments"
#     id = Column(Integer, primary_key=True)
#     name = Column(String(20), nullable=False)
#     manager_id = Column(Integer, ForeignKey("persons.id"))
#     # person = relationship('Person', backref("departments", order_by=id))
#     persons = relationship("Person", backref=backref('departments', order_by=id))
#
#     def __repr__(self):
#         return "id is %s, name is %s, manager_id is %s" % (self.id, self.name, self.manager_id)
#
# p1 = session.query(Person)
# print(p1.all())
#
# d1 = session.query(Department)
# print(d1.all())
#
# r1 = session.query(Person.name, Person.age, Department.name).join(Department, Person.id==Department.manager_id).\
#     filter(Department.name=="Finance")
# print(r1.all())
# # for p, d in session.query(Person, Department).filter(Person.id == Department.manager_id).filter(Department.name=="Finance"):
# #     print(p)
# #     print(d)
# # lucy = Person(name="Lucy", age=20)
# # print(lucy.departments)
# # lucy.departments = [Department(name="logistic"), Department(name="commercial")]
# # print(lucy.departments)
# # print(lucy.departments[1])
# # print(lucy.departments[1].name)
# # session.add(lucy)
# # session.commit()
# # session.commit()
# # p1 = Person()
# # d1 = Department()
# #
# # print(p1.departments)
# # print(d1.persons)
# # p1.departments.append(d1)
# # print(p1.departments)
# # print(d1.persons)
# #
# # r1 = session.query(Person).all()
# # print(r1)
# #
# # r2 = session.query(Department).all()
# # print(r2)
# #
# # r3 = session.query(Department).first()
# # print(r3)
# # print(r3.persons.name)
# #
# # r4 = session.query(Person).first()
# # print(r4)
# # print(r4.departments)
# #
# # # data = Department(name="purchase", manager_id=16)
# # # session.add(data)
# # # session.commit()
# # r5 = session.query(Department)
# # print(r5.all())