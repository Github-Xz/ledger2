from openpyxl import load_workbook

filename = '执法，采集，公示案件文书统计.xlsx'
wb = load_workbook(filename)
ws = wb.active
"""读取指定的sheet表"""
sheet_ranges = wb['01、执法文书当日统计']
sheet_ranges2 = wb['01、执法文书总数统计']

"""获取每日文书数量"""


def get_daywrit(column1, column2):
    row = 6  # 行
    day_writ = {}  # 将每日新增文书名称与数量存入字典

    while True:
        str1 = (column1 + str(row))
        str2 = (column2 + str(row))
        if sheet_ranges[str1].value is None:
            break
        else:
            day_writ[sheet_ranges[str1].value] = sheet_ranges[str2].value  # 将读取到的每日文书数据写入字典
        row += 1
    print('每日文书')
    print(day_writ)
    return day_writ


"""获取总文书数量"""


def get_allwrit(column1, column2):
    row = 4
    all_writ = {}  # 南沙总文书

    while True:
        str1 = (column1 + str(row))
        str2 = (column2 + str(row))
        if sheet_ranges2[str1].value is None:
            break
        else:
            all_writ[sheet_ranges2[str1].value] = sheet_ranges2[str2].value  # 将读取到的总文书数据写入字典
        row += 1
    print('总文书')
    print(all_writ)
    return all_writ


