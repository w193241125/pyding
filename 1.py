#!/usr/local/bin python
# -*- coding: utf-8 -*-
# @Time    : 2018/8/6 19:03
# @Author  : Larwas
# @Site    : 
# @File    : 1.py
# @Software: PyCharm

import xlrd
import xlwt
import types
from datetime import datetime
import time

print('开始读取excel数据... ')
# 背景色数组
bgcolor = []

# 字体颜色数组
ftcolor = []

#加班统计
overtime = []


wb = xlwt.Workbook()
ws = wb.add_sheet('A Test Sheet', cell_overwrite_ok=True)

data = xlrd.open_workbook('1.xlsx')

# sheet_name = '各部门'

# table = data.sheet_by_name(sheet_name)  # 通过工作表名称获取

table = data.sheets()[0]  # 工作表表名

# names = data.sheet_names()    # 返回book中所有工作表的名字

#  data.sheet_loaded(sheet_name or indx)   # 检查某个sheet是否导入完毕

nrows = table.nrows  # 获取该sheet中的有效行数
ncols = table.ncols  # 获取该sheet中的有效列数

for i in range(nrows):
    bgcolor.insert(i, [1])  # 为前三行设置背景颜色元素
    overtime.insert(i, 0)  # 为加班的小伙伴赋初始值 0

# table.row(1)  # 返回由第一行中所有的单元格对象组成的列表

# table.row_slice(1)   # 返回由第一列中所有的单元格对象组成的列表

row_type = table.row_types(1, start_colx=0, end_colx=None)   # 返回由该行中所有单元格的数据类型组成的列表

day_count = len([i for i in row_type if i == 2])  # 获取这个月有多少天

# 获取 迟到次数 所在地方的索引
row_list = table.row_values(2, start_colx=0, end_colx=None)
for_num = row_list.index('迟到次数')  # 得到 迟到次数位置，以确定 for 循环次数
for_overtime = row_list.index('加班天数')  # 得到 加班天数位置，以填充每人的加班时间
for nc in range(ncols):  # 设定每一列宽度
    if 1 < nc < for_num:
        ws.col(nc).width = 256 * 2
    if nc >= for_num:
        ws.col(nc).width = 256 * 4
# 设置内容位置
alignment = xlwt.Alignment()  # Create Alignment
alignment.horz = xlwt.Alignment.HORZ_CENTER  # 水平居中 May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_CENTER  # 是垂直居中 May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED
alignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT  # 自动换行

# 设置边框
borders = xlwt.Borders()
borders.left = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
borders.top = xlwt.Borders.THIN
borders.bottom = xlwt.Borders.THIN



for i in range(nrows):  # 0~4 excel每一行
    rowlist = table.row_values(i, start_colx=0, end_colx=None)  # 返回由第 i 行中所有单元格的数据组成的列表
    if i <= 2:  # 前 3 行是表头信息，不需要加背景色，单独处理
        bgcolor.insert(i, [1])
        if i == 0:  # 第一行合并单元格加粗居中加边框，字体大小 820

            style_a = xlwt.XFStyle()  # Create Style
            style_a.alignment = alignment  # Add Alignment to Style
            style_a.borders = borders
            # 设置字体大小
            font = xlwt.Font()
            font.name = "SimSun"

            font.height = 14 * 20
            font.bold = 'blod'
            style_a.font = font

            # 设置第一行高度
            first_col = ws.col(0)  # xlwt中是行和列都是从0开始计算的
            sec_col = ws.col(1)
            first_col.width = 256 * 3 #  设置列宽为3个 0 的宽度，256为衡量单位，3表示多少个0的宽度
            sec_col.width = 256 * 7

            # 设置字体高度
            for d in range(2, day_count+2):
                ws.col(d).width = 256 * 4
                ws.col(d).height = 256 * 10
                tall_style = xlwt.easyxf('font:height 820;align: wrap on')
                ws.row(d).set_style(tall_style)
            ws.write_merge(0, 0, 0, len(rowlist)-1, rowlist[0], style_a)
        elif i == 1:  # 第二行，
            style_b = xlwt.XFStyle()

            style_b.alignment = alignment
            style_b.borders = borders
            for rc in range(len(rowlist)):
                if rc == 0:
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "SimSun"
                    font.bold = 'blod'
                    font.height = 8 * 20
                    style_b.font = font
                    ws.write_merge(1, 2, 0, 0, rowlist[0], style_b)  # 处理 序号 列
                elif rc == 1:

                    # 设置字体
                    font = xlwt.Font()
                    font.name = "SimSun"
                    font.bold = 'blod'
                    font.height = 10 * 20
                    style_b.font = font
                    ws.write_merge(1, 2, 1, 1, rowlist[rc], style_b)  # 处理 日期/姓名 单元格的合并与样式
                elif rc == (2+day_count):
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "SimSun"
                    font.bold = 'blod'
                    font.height = 5 * 20
                    style_b.font = font
                    ws.write_merge(1, 1, rc, rc+3, rowlist[rc], style_b)  # 处理 罚款/元 单元格的合并与样式
                elif rc == (6+day_count):
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "SimSun"
                    font.bold = 'blod'
                    font.height = 5 * 20
                    style_b.font = font
                    ws.write_merge(1, 1, rc, rc+1, rowlist[rc], style_b)  # 处理 罚款/元 单元格的合并与样式
                elif rc == (8+day_count):
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "SimSun"
                    font.bold = 'blod'
                    font.height = 5 * 20
                    style_b.font = font
                    ws.write_merge(1, 1, rc, rc+3, rowlist[rc], style_b)  # 处理 罚款/元 单元格的合并与样式
                elif rc > (12+day_count):
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "SimSun"
                    font.bold = 'blod'
                    font.height = 5 * 20
                    style_b.font = font
                    ws.write(i, rc, rowlist[rc], style_b)
                elif 2 <= rc < (2+day_count):
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "SimSun"
                    font.bold = 'blod'
                    font.height = 5 * 20
                    style_b.font = font
                    ws.write(i, rc, rowlist[rc], style_b)
                    pass

        elif i == 2:  # 第三行
            for rb in range(len(rowlist)):
                style_c = xlwt.XFStyle()
                # 设置边框
                borders = xlwt.Borders()
                borders.left = xlwt.Borders.THIN
                borders.right = xlwt.Borders.THIN
                borders.top = xlwt.Borders.THIN
                borders.bottom = xlwt.Borders.THIN
                style_c.alignment = alignment
                style_c.borders = borders
                style_c.font = font
                if rb > 1:
                    ws.write(i, rb, rowlist[rb], style_c)
                else:
                    pass

    if i > 2:

        # 处理每一行的每一个单元格
        for k in range(len(rowlist)):  # k 表示每个单元格

            if k == 0:

                style = xlwt.XFStyle()  # Create the Pattern
                # 设置边框
                borders = xlwt.Borders()
                borders.left = xlwt.Borders.THIN
                borders.right = xlwt.Borders.THIN
                borders.top = xlwt.Borders.THIN
                borders.bottom = xlwt.Borders.THIN
                style.borders = borders
                # 设置字体
                font = xlwt.Font()
                font.name = "SimSun"
                font.height = 10 * 20
                style.font = font
                style.alignment = alignment

                ws.write(i, k, rowlist[k], style)
            else:
                rowlist[k] = rowlist[k].replace(' ', '')  # 删除空格

                cell = rowlist[k].split('\n')  # 处理单元格的每个元素 ['08:52  ', '18:54  ', '20:19', '外勤  ']
                # 去除空格
                for j in range(len(cell) - 1):
                    cell[j] = cell[j].replace(' ', '')
                if k <= 2:
                    style_d = xlwt.XFStyle()  # Create the Pattern
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "黑体"
                    font.height = 8 * 20
                    style_d.font = font

                    style_d.alignment = alignment
                    style_d.borders = borders
                    bgcolor[i].insert(k, 1)
                    ws.write(i, k, rowlist[k], style_d)
                if k > 2:
                    # 默认没有背景色
                    bgcolor[i].insert(k, 1)

                    # 处理外勤的小伙伴 标颜色
                    if '外勤' in cell:
                        bgcolor[i].insert(k, 2)  # 红色

                    # 可能迟到的人啊
                    if cell[0] > '09:00' and cell[0] <= '09:30':
                        # 先判断前一天最后打卡时间
                        lastCell = rowlist[k - 1].split('\n')
                        if lastCell[len(lastCell) - 1] != '外勤' and lastCell[len(lastCell) - 1] >= '22:00':
                            pass
                        else:
                            bgcolor[i].insert(k, 3)  # 绿色

                    # 肯定迟到的人啊
                    if cell[0] > '09:30':
                        bgcolor[i].insert(k, 3)  # 绿色

                    # 早退的人啊
                    if cell[len(cell) - 1] != "" and cell[len(cell) - 1] < '18:00':
                        bgcolor[i].insert(k, 4)  # 蓝色

                    # 忘记打卡的人儿啊 或者请假的人儿啊
                    if len(cell) < 2:
                        # 判断是否为周六日
                        if row_list[k] == '六' or row_list[k] == '日':
                            bgcolor[i].insert(k, 1)
                        elif k >= for_num:
                            bgcolor[i].insert(k, 1)  # 不急忘记打卡
                        else:
                            bgcolor[i].insert(k, 5)  # 黄色

                    #  统计加班次数：
                    for x in range(len(cell)):
                        if cell[x] > '21:00' and '外勤' not in cell:
                            overtime[i] += 1


                    pattern = xlwt.Pattern()  # Create the Pattern
                    pattern.pattern = xlwt.Pattern.SOLID_PATTERN  # May be: NO_PATTERN, SOLID_PATTERN, or 0x00 through 0x12
                    style = xlwt.XFStyle()  # Create the Pattern

                    # 设置边框
                    borders = xlwt.Borders()
                    borders.left = xlwt.Borders.THIN
                    borders.right = xlwt.Borders.THIN
                    borders.top = xlwt.Borders.THIN
                    borders.bottom = xlwt.Borders.THIN
                    style.borders = borders
                    # 设置字体
                    font = xlwt.Font()
                    font.name = "黑体"
                    font.height = 6 * 20
                    style.font = font

                    style.alignment = alignment  # Add Alignment to Style
                    # pattern.pattern_fore_colour=? , May be: 8 through 63. 0 = Black, 1 = White, 2 = Red, 3 = Green,
                    # 4 = Blue, 5 = Yellow, 6 = Magenta(紫色), 7 = Cyan:青绿色, 16 = Maroon：褐紫红色,
                    # 17 = Dark Green：暗青色, 18 = Dark Blue：暗蓝色, 19 = Dark Yellow：亮黄色 , almost brown),
                    # 20 = Dark Magenta：暗紫色, 21 = Teal：蓝绿色, 22 = Light Gray, 23 = Dark Gray, the list goes on...

                    if bgcolor[i][k] == 1:
                        # pattern.pattern_fore_colour = 1
                        # style.pattern = pattern  # Add Pattern to Style
                        pass
                    elif bgcolor[i][k] == 2:
                        pattern.pattern_fore_colour = 2
                        style.pattern = pattern
                    elif bgcolor[i][k] == 3:
                        pattern.pattern_fore_colour = 3
                        style.pattern = pattern
                    elif bgcolor[i][k] == 4:
                        pattern.pattern_fore_colour = 4
                        style.pattern = pattern
                    elif bgcolor[i][k] == 5:
                        pattern.pattern_fore_colour = 5
                        style.pattern = pattern
                    elif bgcolor[i][k] == 6:
                        pattern.pattern_fore_colour = 6
                        style.pattern = pattern
                    ws.write(i, k, rowlist[k], style)

        ws.write(i, for_overtime, overtime[i], style)

print('正在生成新文件...')
wb.save('kaoqin.xls')
print('请输入：我是小可爱 退出...')
input("emmm...:")
time.sleep(3)
print('对不起，你不是小可爱')
time.sleep(1)
# table.row_len(4)  # 返回该列的有效单元格长度

