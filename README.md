# **基础学院签到自动统计脚本**

~~写文档真的好烦鸭~~

## 说明：

一个完整的文件结构应该是这样的

```
CBS_Signin
 │   README.md
 │   chart.py
 |   oriinfo.py
 |   models.py
 |   requirements.txt
 |   lessons.xls
 |   2019.xls
 |   signin.db
 |   README.md
 │
 └───import
 |    |   此处留空
 │    |
 |
 └───legacy
 |    |   此处留空
 |    |
 |
 └───output
      │   此处留空
      |
```

### 文件说明：

+ **chart.py**

    用以导出指定时间范围内的信息

+ **oriinfo.py**

    自动解析/import文件夹下的excel表格导入到数据库中

+ **models.py**

    数据库结构定义文件

+ **requeiements.txt**

    使用到的第三方库，可使用 " *pip install -r requirements.txt* " 进行安装

+ **signin.db**

    sqlite3数据库文件

+   **2019.xls**

    新生信息原始数据，表结构如下：

    学号|姓名|性别|学院|专业|班级|出生日期|民族|备注|生源地
    :--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:
    1935010101|小白|女|基础学院|电类|电类01班|20000101|汉族||上海

+ **lesson.xls**

    基础学院所有课程信息，表结构如下：

    学年|学期|校区名称|开课学院|课程代码|课程|课程类别|课程性质|学分|起始结束周|教师|已选人数|上课地点|上课时间|班级组成
    :--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:
    2019-2020|1|校区⑨|理学院|22000762|普通化学B|理论类课程|学科基础课程|2.0|3-18周|小黑|102|二教108|星期五第8-9节{3-18周}|工科实验班

+ **【文件夹】import/**

    用以临时存放将要导入到数据库的文件
    
    #### 表格结构：
    #  <center>签到情况明细</center>
    学年|学期|课程|上课日期|周次|星期|课次|教师工号|教师姓名|学生学号|学生姓名|是否签到|签到时间|手工签到
    :--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:|:--:
    2019-2020|1|模拟电子技术|20190101|0|0|0|0000|小黑|1935010101|小白|1|2019-01-01 00:00:00|0
    ...|...|...|...|...|...|...|...|...|...|...|...|...|...


+ **【文件夹】output/**

    用以存放自动导出的统计文件

+ **【文件夹】legacy/**

    用以存放已经导入到数据库的表格

    所有在import/文件夹下导入完成的表格将被自动移动至此文件夹


## 参考文档：

Peewee库的使用：<http://docs.peewee-orm.com/en/latest/>

openpyxl库的使用：<https://openpyxl.readthedocs.io/en/stable/>