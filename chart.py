from models import *
from oriinfo import dumpstudentinfo,dumplessoninfo,readstudentlist,readlessoninfo

allcourses=[]   #所有课程
allyukecources = [] #预科的课程
allclasses=[]   #所有班级
allyukeclasses=[]   #预科的班级
majorCollegeTranslator={}   #专业学院对照

signinthreshhold = 10 #签到阈值：签到人数大于多少才算这节课考勤了

def output(start,end):
    workbook = Workbook()

    alignment = Alignment(horizontal='center', vertical='center')
    #绿色填充
    style_g = NamedStyle(name='style_g')
    style_g.fill = PatternFill("solid",fgColor='C6E0B4')
    bd = Side(style='thin',color='000000')
    style_g.alignment = alignment
    #蓝色填充
    style_b = NamedStyle(name='style_b')
    style_b.fill = PatternFill("solid", fgColor='BDD7EE')
    bd = Side(style='thin', color='000000')
    style_b.alignment = alignment
    #通用居中
    style_universal = NamedStyle(name='style_universal')
    style_universal.alignment=alignment

    # 工作表1  班级出勤率统计  初始化  样式
    print('正在导出班级出勤率统计...')
    worksheet1 = workbook.active
    worksheet1.title = u'班级出勤率统计'
    worksheet1.sheet_properties.tabColor = '00B050'

    #输出数据
    actualcnt = 0
    datadic={}
    for eachclass in allclasses:
        worksheet1.cell(row=allclasses.index(eachclass)+4, column=1, value=majorCollegeTranslator[eachclass])
        worksheet1.cell(row=allclasses.index(eachclass)+4, column=2, value=eachclass)
        worksheet1.cell(row=allclasses.index(eachclass)+4, column=3, value='')
    col=4
    for eachlesson in allcourses:
        flag=True
        print('\r正在导出第{0}项课程，共{1}项课程'.format(allcourses.index(eachlesson)+1,len(allcourses)),end='')
        for eachclass in allclasses:
            datadic[eachclass] = [0,0]
        alldata = signin.select().where(signin.signincount>signinthreshhold,signin.name==eachlesson,
                                        signin.time>=start,signin.time<=end)
        for each in alldata:
            if each.issignin=='1':
                flag=False
                datadic[each.xzclass][1]+=1
            datadic[each.xzclass][0] += 1
        if flag:
            continue
        actualcnt+=1
        worksheet1.cell(row=2, column=col, value=eachlesson)
        worksheet1.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col + 2)
        worksheet1.cell(row=3, column=col, value=u'人数')
        worksheet1.cell(row=3, column=col + 1, value=u'出勤')
        worksheet1.cell(row=3, column=col + 2, value=u'出勤率')
        for eachclass in allclasses:
            worksheet1.cell(row=allclasses.index(eachclass)+4, column=col, value=str(datadic[eachclass][0]))
            worksheet1.cell(row=allclasses.index(eachclass)+4, column=col + 1, value=str(datadic[eachclass][1]))
            if datadic[eachclass][0]==0:
                worksheet1.cell(row=allclasses.index(eachclass)+4, column=col + 2, value="未考勤")
            else:
                worksheet1.cell(row=allclasses.index(eachclass)+4, column=col + 2, value=str(round(datadic[eachclass][1]/datadic[eachclass][0]*10000)/100)+'%')
        convertrange = ''
        tmp1 = col
        tmp2 = col+2
        tmpstr = ''
        while tmp1:
            tmp1 -= 1
            tmpstr += chr(tmp1 % 26 + ord('A'))
            tmp1 //= 26
        convertrange += tmpstr[::-1]
        convertrange += ':'
        tmpstr = ''
        while tmp2:
            tmp2 -= 1
            tmpstr += chr(tmp2 % 26 + ord('A'))
            tmp2 //= 26
        convertrange += tmpstr[::-1]
        if col % 2:
            for each_col in worksheet1[convertrange]:
                for cell in each_col:
                    cell.style = style_g
        else:
            for each_col in worksheet1[convertrange]:
                for cell in each_col:
                    cell.style = style_b
        col+=3
    for each_col in worksheet1['A:C']:
        for cell in each_col:
            cell.style = style_universal
    worksheet1.column_dimensions['A'].width = 25
    worksheet1.column_dimensions['B'].width = 15
    print('')
    # 表头制作
    worksheet1.cell(row=1, column=1, value=u'签到情况明细')
    worksheet1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=(actualcnt + 1) * 3)
    worksheet1.cell(row=2, column=1, value='')
    worksheet1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)
    worksheet1['A3'].value = u'学院'
    worksheet1['B3'].value = u'班级'
    worksheet1['C3'].value = u'班级人数'

    # 工作表2  预科
    print('正在导出预科缺勤统计...')
    worksheet1 = workbook.create_sheet('预科缺勤统计')
    worksheet1.sheet_properties.tabColor = '00B050'

    worksheet1.cell(row=1, column=1, value=u'预科签到情况明细')
    worksheet1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=(len(allyukecources) + 1) * 3)
    worksheet1.cell(row=2, column=1, value='')
    worksheet1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)
    worksheet1['A3'].value = u'学院'
    worksheet1['B3'].value = u'班级'
    worksheet1['C3'].value = u'班级人数'

    i = 0
    for each in allyukecources:
        worksheet1.cell(row=2, column=4 + i * 3, value=each)
        worksheet1.merge_cells(start_row=2, start_column=4 + i * 3, end_row=2, end_column=6 + i * 3)
        worksheet1.cell(row=3, column=4 + i * 3, value=u'人数')
        worksheet1.cell(row=3, column=5 + i * 3, value=u'出勤')
        worksheet1.cell(row=3, column=6 + i * 3, value=u'出勤率')
        i += 1

    # 输出数据
    datadic = {}
    for eachclass in allyukeclasses:
        worksheet1.cell(row=allyukeclasses.index(eachclass) + 4, column=1, value=majorCollegeTranslator[eachclass])
        worksheet1.cell(row=allyukeclasses.index(eachclass) + 4, column=2, value=eachclass)
        worksheet1.cell(row=allyukeclasses.index(eachclass) + 4, column=3, value='')
    col = 4
    for eachlesson in allyukecources:
        print('\r正在导出第{0}项课程，共{1}项课程'.format(allyukecources.index(eachlesson) + 1, len(allyukecources)), end='')
        for eachclass in allyukeclasses:
            datadic[eachclass] = [0, 0]
        alldata = signin.select().where(signin.signincount > signinthreshhold, signin.name == eachlesson,
                                        signin.time >= start, signin.time <= end)
        for each in alldata:
            if each.issignin == '1':
                datadic[each.xzclass][1] += 1
            datadic[each.xzclass][0] += 1
        for eachclass in allyukeclasses:
            worksheet1.cell(row=allyukeclasses.index(eachclass) + 4, column=col, value=str(datadic[eachclass][0]))
            worksheet1.cell(row=allyukeclasses.index(eachclass) + 4, column=col + 1, value=str(datadic[eachclass][1]))
            if datadic[eachclass][0] == 0:
                worksheet1.cell(row=allyukeclasses.index(eachclass) + 4, column=col + 2, value="未考勤")
            else:
                worksheet1.cell(row=allyukeclasses.index(eachclass) + 4, column=col + 2,
                                value=str(round(datadic[eachclass][1] / datadic[eachclass][0] * 10000) / 100) + '%')
        convertrange = ''
        tmp1 = col
        tmp2 = col + 2
        tmpstr = ''
        while tmp1:
            tmp1 -= 1
            tmpstr += chr(tmp1 % 26 + ord('A'))
            tmp1 //= 26
        convertrange += tmpstr[::-1]
        convertrange += ':'
        tmpstr = ''
        while tmp2:
            tmp2 -= 1
            tmpstr += chr(tmp2 % 26 + ord('A'))
            tmp2 //= 26
        convertrange += tmpstr[::-1]
        if col % 2:
            for each_col in worksheet1[convertrange]:
                for cell in each_col:
                    cell.style = style_g
        else:
            for each_col in worksheet1[convertrange]:
                for cell in each_col:
                    cell.style = style_b
        col += 3
    for each_col in worksheet1['A:C']:
        for cell in each_col:
            cell.style = style_universal
    worksheet1.column_dimensions['A'].width = 25
    worksheet1.column_dimensions['B'].width = 15
    print('')

    # 工作表3  缺勤名单  初始化
    print('正在导出缺勤名单...')
    worksheet1 = workbook.create_sheet('缺勤名单')
    worksheet1.sheet_properties.tabColor = 'FF0000'
    alldata = signin.select().where(signin.issignin == 0, signin.signincount > 10, signin.time >= start,
                                    signin.time <= end).order_by(signin.studentid)

    #表头制作
    worksheet1.cell(row=1, column=1, value=u'学院')
    worksheet1.cell(row=1, column=2, value=u'班级')
    worksheet1.cell(row=1, column=3, value=u'学号')
    worksheet1.cell(row=1, column=4, value=u'姓名')
    worksheet1.cell(row=1, column=5, value=u'日期')
    worksheet1.cell(row=1, column=6, value=u'缺勤课程')

    #输出数据
    i = 2
    for each in alldata:
        worksheet1.cell(row=i, column=1, value=each.college)
        worksheet1.cell(row=i, column=2, value=each.xzclass)
        worksheet1.cell(row=i, column=3, value=each.studentid)
        worksheet1.cell(row=i, column=4, value=each.studentname)
        worksheet1.cell(row=i, column=5, value=each.time)
        worksheet1.cell(row=i, column=6, value=each.name)
        i += 1
    worksheet1.column_dimensions['A'].width = 25
    worksheet1.column_dimensions['B'].width = 15
    worksheet1.column_dimensions['C'].width = 13
    worksheet1.column_dimensions['D'].width = 30
    worksheet1.column_dimensions['E'].width = 10
    worksheet1.column_dimensions['F'].width = 40

    for each_col in worksheet1['A:F']:
        for cell in each_col:
            cell.style = style_universal

    # 工作表4  没有考勤的老师  初始化
    print('正在导出未考勤课程清单...')
    worksheet1 = workbook.create_sheet('未考勤课程')
    worksheet1.sheet_properties.tabColor = 'FF0000'
    alldata = classinfo.select().where(classinfo.classsignin <= 10, classinfo.time >= start,
                                    classinfo.time <= end)

    # 表头制作
    worksheet1.cell(row=1, column=1, value=u'课程名称')
    worksheet1.cell(row=1, column=2, value=u'日期')
    worksheet1.cell(row=1, column=3, value=u'周数')
    worksheet1.cell(row=1, column=4, value=u'星期')
    worksheet1.cell(row=1, column=5, value=u'课次')
    worksheet1.cell(row=1, column=6, value=u'教师工号')
    worksheet1.cell(row=1, column=7, value=u'教师姓名')

    # 输出数据
    i = 2
    for each in alldata:
        worksheet1.cell(row=i, column=1, value=each.Class)
        worksheet1.cell(row=i, column=2, value=each.time)
        worksheet1.cell(row=i, column=3, value=each.week)
        worksheet1.cell(row=i, column=4, value=each.day)
        worksheet1.cell(row=i, column=5, value=each.no)
        worksheet1.cell(row=i, column=6, value=each.teacherno)
        worksheet1.cell(row=i, column=7, value=each.teachername)
        i += 1
    worksheet1.column_dimensions['A'].width = 25
    worksheet1.column_dimensions['B'].width = 15
    worksheet1.column_dimensions['C'].width = 10
    worksheet1.column_dimensions['D'].width = 10
    worksheet1.column_dimensions['E'].width = 10
    worksheet1.column_dimensions['F'].width = 10
    worksheet1.column_dimensions['G'].width = 20

    for each_col in worksheet1['A:G']:
        for cell in each_col:
            cell.style = style_universal

    #保存表
    if start!=end:
        filename = 'output/'+str(start)+'-'+str(end)+'签到统计.xlsx'
    else:
        filename = 'output/'+str(start)+'签到统计.xlsx'
    workbook.save(filename)

if __name__ == '__main__':
    db.connect()

    # studentinfo,studentlist = readstudentlist(xlrd.open_workbook('2019.xls'))   #从excel读信息
    studentinfo, studentlist = dumpstudentinfo()  # 从数据库读信息

    # lessoninfo = readlessoninfo(xlrd.open_workbook('lesson.xls'))
    lessoninfo = dumplessoninfo()

    for each in lessoninfo:
        if lessoninfo[each]=='学科基础课程':
            allcourses.append(each)
        elif lessoninfo[each]=='综合素质必修':
            allyukecources.append(each)
    for each in studentinfo:
        if studentinfo[each]['Class'] not in allclasses and studentinfo[each]['Class'] not in allyukeclasses:
            majorCollegeTranslator[studentinfo[each]['Class']]=studentinfo[each]['college']
            if studentinfo[each]['major'] == '少数民族预科生':
                allyukeclasses.append(studentinfo[each]['Class'])
            else:
                allclasses.append(studentinfo[each]['Class'])

    startdate = input("输入开始日期(输入格式:20191016)[包含]：")
    enddate = input("输入结束日期(输入格式:20191016)[包含]：")
    output(startdate,enddate)