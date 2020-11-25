import xlrd
from xlrd import xldate_as_tuple
import datetime
from prettytable import PrettyTable
import xlsxwriter

def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook(r'./kaoqin202011.xlsx')
    # 得到第一个sheet
    sheet = workbook.sheet_by_index(0)
    dict = {}
    for i in range(1, sheet.nrows):
            # 读取员工序号
            id = sheet.cell(i, 1).value
            if id not in dict:
                dict[id] = {}
            # 读取打卡时间
            date = datetime.datetime(*xldate_as_tuple(sheet.cell(i, 2).value, 0))
            dt = date.strftime('%Y%m%d')
            dtime = date.strftime('%H%M%S')
            if dt not in dict[id]:
                dict[id][dt] = []
            dict[id][dt].append(dtime)
    return dict

def getDateList(start_date, end_date):
    date_list = []
    start_date = datetime.datetime.strptime(start_date, '%Y-%m-%d')
    end_date = datetime.datetime.strptime(end_date, '%Y-%m-%d')
    date_list.append(start_date.strftime('%Y%m%d'))
    while start_date < end_date:
        start_date += datetime.timedelta(days=1)
        date_list.append(start_date.strftime('%Y%m%d'))
    return date_list

def statistic(data,date):
    result = {}
    for key in data:
        tlist = []
        row = data[key]
        # 循环天数
        for d in date:
            tip = ""
            atip = ""
            stip = ""
            etip = ""
            if d not in row:
                tip = "缺卡"
            else:
                # 取最早的时间
                start_time = min(row[d])
                end_time = max(row[d])
                # 如果只有一个卡打时间 则判断
                if start_time > '083000':
                    stip = '迟到--打卡时间为' + start_time
                if end_time < '170000':
                    etip = tip + "\r\n早退--打卡时间为" + end_time
            tip = tip + stip + etip
            if tip == "":
                tip = "正常"
            tlist.append(tip)
        result[key] = tlist
    return result

def save(data,head):
    workbook = xlsxwriter.Workbook("result20201025-20201124.xlsx")
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    ItemStyle = workbook.add_format({
        'align': 'center',
        'text_wrap': 1,
        'valign': 'vcenter'
    })
    # 表头
    head.insert(0, '考勤号码')
    worksheet.write_row('A1', head, ItemStyle)
    worksheet.set_column(0, len(head), 20)
    i = 2
    for key in data:
        row = data[key]
        row.insert(0, key)
        worksheet.write_row('A' + str(i), row, ItemStyle)
        i = i+1
    workbook.close()





