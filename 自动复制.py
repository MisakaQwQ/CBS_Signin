import os
import time
import shutil
dir1 = r'C:\Users\charq\Downloads'
dir2 = r'C:\Users\charq\Desktop\workspace\Signin\import'

day=1
week=15
while True:
    a=os.listdir(dir1)
    if r'签到情况明细.xls' in a:
        newname = r'\第'+str(week)+'周'+str(day)+'.xls'
        shutil.move(dir1+r'\签到情况明细.xls',dir2+newname)
        print('已复制{0}'.format(newname))
        day+=1
        if day >=6:
            day=1
            week+=1
        time.sleep(3)
