import calendar

def weekdayName(dayName):
    if dayName == 0:
        dayName = '周一'
    elif dayName == 1:
        dayName = '周二'
    elif dayName == 2:
        dayName = '周三'
    elif dayName == 3:
        dayName = '周四'
    elif dayName == 4:
        dayName = '周五'
    elif dayName == 5:
        dayName = '周六'
    else:
        dayName = '周日'
    return dayName


dayName = calendar.weekday(2018,8,1)
a=calendar.monthrange(2002,1)
print(type(a))

year = 2018
month = 8
days=calendar.monthrange(2018,8)    #该月共几天


for i in range(days[0]-1,days[1]+1):
    cal = calendar.weekday(int(year),int(month),i) #该月份第一天是周几
    dayName=weekdayName(cal)
    print(dayName)
    #ws.cell(row=2, column=2+i).value=dayName