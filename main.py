# This is a sample Python script.
import readExcel
if __name__ == '__main__':
    data = readExcel.read_excel()
    alldate = readExcel.getDateList('2020-10-25', '2020-11-24')
    result = readExcel.statistic(data, alldate)
    readExcel.save(result, alldate)
