#encoding=utf-8
from openpyxl import Workbook
from openpyxl.styles import Font,colors
import datetime
import calendar

#函数部分
#数字转周几
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



year =datetime.datetime.now().strftime("%Y")
month = datetime.datetime.now().strftime("%m")
days=calendar.monthrange(int(year),int(month))    #该月共几天



wb = Workbook()
ws=wb.active
#合并单元格
ws.merge_cells('B1:E1')
ws.merge_cells('A4:A7')
ws.merge_cells('A8:A10')
ws.merge_cells('A11:A12')
ws.merge_cells('A14:A15')
ws.merge_cells('A18:A23')
ws.merge_cells('A24:A28')
ws.merge_cells('A29:A30')
ws.merge_cells('A34:N34')
ws.merge_cells('A35:N35')

#填内容
#表单名
ws['B1']='%s年%s月份住宿登记表'%(year,month)
#金浦花园
ws['A2']='金浦花园'
ws['A4']='纽约'
ws['A8']='巴黎'
ws['A11']='伦敦'
ws['A13']='米兰'
ws['A14']='厅两人间'
#床位号
ws['B4']='四人间1号床'
ws['B5']='四人间2号床'
ws['B6']='四人间3号床'
ws['B7']='四人间4号床'
ws['B8']='三人间1号床'
ws['B9']='三人间2号床'
ws['B10']='三人间3号床'
ws['B11']='两人间2号床'
ws['B12']='两人间3号床'
ws['B13']='单人间'
ws['B14']='沙发床'
ws['B15']='单床'

#微山三村
ws['A16']='微山三村'
ws['A18']='伦敦（男生）'
ws['A24']='上海（女生）'
ws['A29']='巴黎'
ws['A31']='米兰'
ws['B18']='WL1上'
ws['B19']='WL1下'
ws['B20']='WL2上'
ws['B21']='WL2下'
ws['B22']='WL3上'
ws['B23']='WL3下'
ws['B24']='WS1上'
ws['B25']='WS1下'
ws['B26']='WS2上'
ws['B27']='WS2下'
ws['B28']='WS3'
ws['B29']='WP1'
ws['B30']='WP2'
ws['B31']='WM'
#周几及日期
for i in range(days[0]-1,days[1]+1):
    cal = calendar.weekday(int(year),int(month),i) #该月份第一天是周几
    dayName=weekdayName(cal)
    ws.cell(row=2, column=2+i).value=dayName    #周几
    ws.cell(row=3, column=2+i).value=i      #日期
#共用
ws['A3']=ws['A17']='房间名'
ws['B2']=ws['B16']='日期'
ws['B3']=ws['B17']='床号'
ws['A2'].font=ws['A16'].font=Font(color=colors.RED,bold=True)#红色加粗字体
ws['B1'].font=Font(color=colors.RED)#红色字体


for i in range(2,32):               #B2-B31字体加粗
        ws['B%s'%i].font=Font(bold=True)


#备注
ws['A34']='备注：金浦花园地址南泉路1261弄，靠近兰村路，电梯房,价格米兰房间100元/天，伦敦房间一床70元/天，巴黎房间一床80元/天，纽约房间一床75元/天'
ws['A35']='      微山三村地址微山路浦明路口，楼梯房6楼，要订房请确认能爬楼梯'
filename = '住宿%s月份.xlsx'%month
wb.save(filename=filename)

