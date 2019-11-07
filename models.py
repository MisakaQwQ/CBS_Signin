from peewee import *
import os
import re
import xlrd
import xlwt
from xlutils.copy import copy
from xlrd import open_workbook
from datetime import date,datetime
from models import *
from playhouse.shortcuts import model_to_dict,dict_to_model
from openpyxl import Workbook
from openpyxl.styles import Font,NamedStyle,Border,Side,PatternFill,Alignment
import shutil
db = SqliteDatabase('signin.db')

class lesson(Model):    #课程信息
    year = CharField(null = True)   #学年
    term = CharField(null = True)   #学期
    school = CharField(null = True) #校区
    college = CharField(null = True)    #开课学院
    code = CharField(null = True)   #课程代码
    name = CharField(null = True)   #课程名称
    sort = CharField(null = True)   #课程类别
    property = CharField(null = True)   #课程性质
    gpd = CharField(null = True)    #学分
    duration = CharField(null = True)   #起始结束周
    teacher = CharField(null = True)    #教师
    capacity = CharField(null = True)   #已选人数
    location = CharField(null = True)   #上课地点
    time = CharField(null = True)   #上课时间
    comp = CharField(null = True)   #教学班组成

    class Meta:
        database = db

class info(Model):  #学生信息
    id = CharField(unique=True, null=False, primary_key=True)   #学号
    name = CharField(null = True)   #姓名
    sex = CharField(null = True)    #性别
    college = CharField(null = True)    #学院
    major = CharField(null = True)  #专业
    Class = CharField(null = True)  #班级
    birthday = CharField(null = True)   #出生日期
    mz = CharField(null = True) #民族
    From = CharField(null = True)   #生源地
    ps = CharField(null = True) #备注
    class Meta:
        database = db

class signin(Model):
    year = CharField(null = True)   #学年
    term = CharField(null = True)   #学期
    name = CharField(null = True)   #课程名称
    time = IntegerField(null = True)   #上课日期
    week = IntegerField(null = True)    #周次
    day = CharField(null = True)    #星期
    no = CharField(null = True) #课次
    teacherid = CharField(null = True)  #教师工号
    teachername = CharField(null = True)    #教师姓名
    college = CharField(null=True)  #学生学院
    major = CharField(null=True)    #学生专业
    xzclass = CharField(null=True)  #学生班级
    studentid = CharField(null = True)  #学生学号
    studentname = CharField(null = True)    #学生姓名
    issignin = CharField(null = True)   #是否签到
    signintime = CharField(null = True) #签到时间
    manualsignin = CharField(null = True)   #手工签到
    classid = IntegerField(null = True) #课程唯一识别号
    signincount = IntegerField(null = True) #课程总签到人数
    From = CharField(null = True)   #来源文件
    class Meta:
        database = db

class classinfo(Model): #暂定 可能可以删除了
    id = IntegerField(unique=True, null=False, primary_key=True)    #课程唯一识别号
    year = CharField(null=True)
    term = CharField(null=True)
    Class = CharField(null=True)
    time = IntegerField(null=True)
    week = IntegerField(null=True)
    day = CharField(null = True)
    no = CharField(null=True)
    teacherno = CharField(null=True)
    teachername = CharField(null=True)
    classsignin = IntegerField(null=True)
    totalsignin = IntegerField(null=True)
    class Meta:
        database = db
