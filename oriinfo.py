from models import *

def readstudentlist(workbook):
    print("正在读取学生信息...")
    #读excel
    allstudent = {}
    studentlist=[]
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    LineKeys = ['id','name','sex','college','major','Class','birthday','mz','ps','From']
    for i in range(1,nrows):
        alineval=[]
        for j in range(0,ncols):
            alineval.append(sheet.cell_value(i,j))
        alineval = dict((key,value) for key,value in zip(LineKeys,alineval))
        allstudent[alineval['id']]=alineval
        #print(alineval)
    allstudent = dict(sorted(allstudent.items(), key=lambda item: item[0], reverse=False))
    # print(allstudent)
    #写数据库
    print("导入学生信息中...")
    data = []
    for each in allstudent:
        data.append(allstudent[each])
        studentlist.append(str(each))
    with db.atomic():
        for idx in range(0, len(data), 50):
            rows = data[idx:idx + 50]
            info.insert_many(rows).execute()
    # print(data)
    print("导入完成")
    return allstudent,studentlist

def dumpstudentinfo():
    q = info.select()
    allstudent={}
    studentlist=[]
    for each in q:
        allstudent[str(each)]=model_to_dict(each)
        studentlist.append(str(each))
    return allstudent,studentlist

def readlessoninfo(workbook):
    alllesson=[]
    lessondict={}
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    LineKeys = ['year', 'term', 'school', 'college', 'code', 'name', 'sort', 'property', 'gpd', 'duration',
                'teacher', 'capacity', 'location', 'time', 'comp']
    for i in range(1, nrows):
        alineval = []
        for j in range(0, ncols):
            alineval.append(sheet.cell_value(i, j))
        alineval = dict((key, value) for key, value in zip(LineKeys, alineval))
        alllesson.append(alineval)
        lessondict[alineval['name']]=alineval['property']
    with db.atomic():
        for idx in range(0, len(alllesson), 50):
            rows = alllesson[idx:idx + 50]
            lesson.insert_many(rows).execute()
    return lessondict

def dumplessoninfo():
    q = lesson.select()
    alllesson={}
    for each in q:
        alllesson[each.name]=each.property
    return alllesson

def process(workbook,studentinfo,studentlist,f):
    data=[]
    classdata={}
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    LineKeys = ['year','term','name','time','week','day','no','teacherid','teachername',
                'studentid','studentname','issignin','signintime','manualsignin']
    for i in range(2, nrows):
        print("\r正在处理第{0}行，共{1}行，进度条{2}%".format(i-1,nrows-2,round(int((i-1)/(nrows-2)*10000))/100),end='')
        alineval = []
        for j in range(0, ncols):
            alineval.append(sheet.cell_value(i, j))
            try:
                alineval[-1]=int(alineval[-1])
            except:
                pass
        alineval = dict((key, value) for key, value in zip(LineKeys, alineval))
        alineval['studentid']=str(alineval['studentid'])
        if alineval['studentid'] not in studentlist:
            continue
        alineval['college']=studentinfo[alineval['studentid']]['college']
        alineval['major'] = studentinfo[alineval['studentid']]['major']
        alineval['xzclass'] = studentinfo[alineval['studentid']]['Class']
        alineval['classid'] = int(str(alineval['term'])+str(alineval['time'])+str(alineval['no'])+str(alineval['teacherid']))
        alineval['signincount'] = 0
        alineval['From'] = f
        data.append(alineval)
        if alineval['classid'] not in classdata:
            classrow = {'id':alineval['classid'], 'year':alineval['year'],'term':alineval['term'],
                        'Class':alineval['name'],'time':alineval['time'],'week':alineval['week'],
                        'day':alineval['day'],'no':alineval['no'],'teacherno':alineval['teacherid'],
                       'classsignin':0,'totalsignin':0,'teachername':alineval['teachername']}
            classdata[alineval['classid']]=classrow
        if alineval['issignin']==1:
            classdata[alineval['classid']]['classsignin']+=1
        classdata[alineval['classid']]['totalsignin']+=1
    print('')
    for each in data:
        each['signincount']=classdata[each['classid']]['classsignin']
    with db.atomic():
        for idx in range(0, len(data), 25):
            rows = data[idx:idx + 25]
            signin.insert_many(rows).execute()
    classdump=[]
    for each in classdata:
        classdump.append(classdata[each])
    with db.atomic():
        for idx in range(0, len(classdump), 25):
            rows = classdump[idx:idx + 25]
            classinfo.insert_many(rows).execute()

if __name__ == '__main__':
    try:
        db.connect()
        db.create_tables([lesson,info , signin, classinfo])
    except:
        db.create_tables([lesson,info , signin, classinfo])

    # studentinfo,studentlist = readstudentlist(xlrd.open_workbook('2019.xls'))   #从excel读信息
    studentinfo,studentlist = dumpstudentinfo()  #从数据库读信息

    # lessoninfo = readlessoninfo(xlrd.open_workbook('lesson.xls'))
    lessoninfo = dumplessoninfo()

    dir = 'import/'
    files = os.listdir(dir)
    for f in files:
        print("正在导入...共{0}个文件，第{1}个文件".format(len(files),files.index(f)+1))
        if os.path.splitext(f)[1] not in ['.xls', 'xlsx']:
            continue
        c = signin.select().where(signin.From == f)
        if len(c) > 0:
            continue
        process(xlrd.open_workbook(dir+f),studentinfo,studentlist,f)
        shutil.move(dir+f,'legacy/'+f)  #处理完就挪走
