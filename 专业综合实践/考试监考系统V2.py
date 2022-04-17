"""
 -*- coding: utf-8 -*-

 @Time : 2021/11/6 15:37

 @Author : jagger

 @File : 考试监考系统.py

 @Software: PyCharm 

 @contact: 252587809@qq.com

 -*- 功能说明 -*-

"""
import webbrowser
from tkinter import *
from tkinter import messagebox

import xlrd
import xlwt
from PIL import Image, ImageTk  # PIL图像处理库
from xlutils.copy import copy

resoultFile = '2021齐大期末考试监考名单.xls'
sourceFile = '监考人员-科目-教师名单.xls'

# 每个列表的第0行存的是提示信息
# 内容是字典对象
teacherList = []  # 监考人员名单 74个
subjectList = []  # 考试科目 53门 规格为30、|60、70、|120、200|
classroomList = []  # 空闲教室 规格为30、90、200
resoultList = []  # 监考结果集 存放的是字典对象

# 监考时间表
time = ['2022/01/03 8:30-11:00', '2022/01/03 14:30-16:00',
        '2022/01/04 8:30-11:00', '2022/01/04 14:30-16:00',
        '2022/01/05 8:30-11:00', '2022/01/05 14:30-16:00',
        '2022/01/06 8:30-11:00', '2022/01/06 14:30-16:00',
        '2022/01/07 8:30-11:00', '2022/01/07 14:30-16:00',
        '2022/01/10 8:30-11:00', '2022/01/10 14:30-16:00']


def createExcel():
    '''
    用于创建和初始化excel表格的风格和基本信息
    样式：
    r1        2021齐大期末考试监考名单
    r2  监考日期  考试科目  考试人数   监考教室  教室容纳人数  监考老师
    @return:
    @rtype:
    '''

    # 创建表格
    workBook = xlwt.Workbook(encoding='utf-8')
    # cell_overwrite_ok=True 允许重载
    workBook_1 = workBook.add_sheet('监考名单', cell_overwrite_ok=True)
    # 合并单元格
    workBook_1.write_merge(0, 1, 0, 5, '2021齐大期末考试监考名单', get_style_01(True))
    # "merge(self, r1, r2, c1, c2, style=Style.default_style):"
    # 形参说明：r1 起始行，r2 合并终止行，c1 起始列 c2 合并终止列
    workBook_1.write(2, 0, '监考日期', get_style_02())
    workBook_1.write(2, 1, '考试科目', get_style_02())
    workBook_1.write(2, 2, '考试人数', get_style_02())
    workBook_1.write(2, 3, '监考教室', get_style_02())
    workBook_1.write(2, 4, '教室容纳人数', get_style_02())
    workBook_1.write(2, 5, '监考老师', get_style_02())

    # 设置单元格宽度
    workBook_1.col(0).width = 400 * 20
    workBook_1.col(1).width = 350 * 20
    workBook_1.col(2).width = 120 * 20
    workBook_1.col(3).width = 120 * 20
    workBook_1.col(4).width = 200 * 20
    workBook_1.col(5).width = 350 * 20
    # workBook_1.write(3,0,'我是第三行',get_style_02())

    # 保存
    workBook.save(resoultFile)

    # # 写入上午日期
    # def write_data_morning():
    #     morning = ['2022/01/03 8:30-11:00', '2022/01/04 8:30-11:00', '2022/01/05 8:30-11:00', '2022/01/06 8:30-11:00',
    #                '2022/01/07 8:30-11:00', '2022/01/10 8:30-11:00']
    #     old_excel = xlrd.open_workbook(resoultFile, formatting_info=True)
    #     new_excel = copy(old_excel)
    #     rws = new_excel.get_sheet(0)
    #     t = 0
    #     for i in range(3, 55, 9):
    #         for j in range(5):
    #             rws.write(i + j, 0, morning[t], get_style_02())
    #         t = t + 1
    #
    #     new_excel.save(resoultFile)
    #
    #
    # # 写入下午日期
    # def write_data_afternoon():
    #     old_excel = xlrd.open_workbook(resoultFile, formatting_info=True)
    #     new_excel = copy(old_excel)
    #     rws = new_excel.get_sheet(0)
    #     morning = ['2022/01/03 14:30-16:00', '2022/01/04 14:30-16:00', '2022/01/05 14:30-16:00', '2022/01/06 14:30-16:00',
    #                '2022/01/07 8:30-11:00', '2022/01/10 14:30-16:00']
    #     t = 0
    #     for i in range(8, 47, 9):
    #         for j in range(4):
    #             rws.write(i + j, 0, morning[t], get_style_02())
    #         t = t + 1
    #
    #     for i in range(53, 56):
    #         rws.write(i, 0, morning[5], get_style_02())
    #     new_excel.save(resoultFile)

    # 写入考试科目以及考试科目人数


def reader():
    '''
    读取信息文件
    Returns:

    '''
    work_book = xlrd.open_workbook(sourceFile)
    teacher_table = work_book.sheet_by_name('监考人员名单')
    subject_table = work_book.sheet_by_name('考试科目')
    classroom_table = work_book.sheet_by_name('空闲教室')
    for i in teacher_table:
        __teacher = {'id': i[0].value, 'dept': i[1].value, 'name': i[2].value, 'sex': i[3].value}
        teacherList.append(__teacher)
    for i in subject_table:
        __subject = {'id': i[0].value, 'subject': i[1].value, 'amount': i[2].value}
        subjectList.append(__subject)
    for i in classroom_table:
        __classroom = {'id': i[0].value, 'classroom': i[1].value, 'amount': i[2].value}
        classroomList.append(__classroom)
    # print('teacherList')
    # print(teacherList)
    # print('subjectList')
    # print(subjectList)
    # print('classroomList')
    # print(classroomList)


def bubbleSortR(sourceList):
    '''
    升序的冒泡排序的python实现
    @param sourceList:传入一个列表
    @type sourceList:列表
    @return: arr 经过排序后的新list 从上到下越来越大
    @rtype:
    '''

    arr = sourceList
    n = len(arr)

    # 遍历所有数组元素
    for i in range(n):

        # Last i elements are already in place
        for j in range(0, n - i - 1):

            if arr[j]['amount'] > arr[j + 1]['amount']:
                arr[j], arr[j + 1] = arr[j + 1], arr[j]
    # print(arr)
    return arr


def bubbleSortD(sourceList):
    '''
    降序的冒泡排序的python实现
    @param sourceList:传入一个列表
    @type sourceList:列表
    @return: arr 经过排序后的新list 从上到下越来越小
    @rtype:
    '''

    arr = sourceList
    n = len(arr)

    # 遍历所有数组元素
    for i in range(n):

        # Last i elements are already in place
        for j in range(0, n - i - 1):

            if arr[j]['amount'] < arr[j + 1]['amount']:
                arr[j], arr[j + 1] = arr[j + 1], arr[j]
    # print(arr)
    return arr


def quickSort(arr, left=None, right=None):
    '''
    快速排序的递归实现
    Parameters
    ----------
    arr : 待排序数组
    left : 首个元素
    right : 最后一个元素

    Returns
    -------
    arr 结果数组
    '''
    left = 0 if not isinstance(left, (int, float)) else left
    right = len(arr) - 1 if not isinstance(right, (int, float)) else right
    if left < right:
        partitionIndex = partition(arr, left, right)
        # 根据枢轴划分为两个小的无序区间
        quickSort(arr, left, partitionIndex - 1)
        quickSort(arr, partitionIndex + 1, right)
    return arr


def partition(arr, left, right):
    '''
    一趟划分
    Parameters
    ----------
    arr : 待排序数组
    left : 首个元素
    right : 最后一个元素

    Returns
    -------

    '''
    pivot = left  # 枢轴
    index = pivot + 1
    i = index
    while i <= right:
        if arr[i] < arr[pivot]:
            swap(arr, i, index)
            index += 1
        i += 1
    swap(arr, pivot, index - 1)  # 讲枢轴与最后一个比他小的元素换位置，此时枢轴左边的元素都比他小
    return index - 1  # 返回此时枢轴位置


def swap(arr, i, j):
    '''
    交换函数
    Parameters
    ----------
    arr : 目标数组
    i : 元素下标
    j : 元素下标

    Returns
    -------

    '''
    arr[i], arr[j] = arr[j], arr[i]


def allocate():
    '''
    监考分配逻辑的具体实现
    @return:
    @rtype:
    '''
    _subjectList = bubbleSortD(subjectList)
    _classroomList = bubbleSortD(classroomList)
    _teacherListMan = []
    _teacherListWoman = []

    for i in teacherList:
        if i['sex'] == '女':
            _teacherListWoman.append(i['name'])
        else:
            _teacherListMan.append(i['name'])
    k, l, m = 0, 0, 0
    while len(subjectList) > 0:
        for i in _classroomList:
            flag = 0  # 标志是否开始重新分配教室
            if len(subjectList) == 0:
                break

            resoultDic = {}  # {监考日期  考试科目  考试人数   监考教室  教室容纳人数  监考老师}
            # 0、填充教室
            resoultDic['监考教室'] = i['classroom']
            resoultDic['教室容纳人数'] = i['amount']

            # 分配考试时间
            resoultDic['监考日期'] = time[m]

            # 2、为每个教室分配老师
            resoultDic['监考老师'] = ' ' + _teacherListWoman[k] + ' ' + _teacherListMan[l]

            # 1、将每个教室分配好考试科目
            for j in _subjectList:
                if j['amount'] <= i['amount']:  # 当考试人数小于教室容纳人数
                    resoultDic['考试科目'] = j['subject']
                    resoultDic['考试人数'] = j['amount']
                    _subjectList.remove(j)
                else:
                    flag = 1
                break
                # 重新开始遍历教室
            if flag == 0:
                # 3、将分配结果放到列表
                m += 1
                m %= len(time)
                k += 1
                k %= len(_teacherListWoman)
                l += 1
                l %= len(_teacherListMan)
                resoultList.append(resoultDic)
            elif flag == 1:
                break
    # print('ok')


def writer():
    '''
    将分配结果写入文件
    @return:
    @rtype:
    '''
    # work_book = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
    source_excel = xlrd.open_workbook(resoultFile, formatting_info=True)
    work_book = copy(source_excel)
    table = work_book.get_sheet(0)
    j = 3
    for i in resoultList:
        table.write(j, 0, i['监考日期'])
        table.write(j, 1, i['考试科目'])
        table.write(j, 2, i['考试人数'])
        table.write(j, 3, i['监考教室'])
        table.write(j, 4, i['教室容纳人数'])
        table.write(j, 5, i['监考老师'])
        j += 1

    work_book.save(resoultFile)  # 保存工作簿


# 打开Excel表格
def open_word():
    webbrowser.open('2021齐大期末考试监考名单.xls')


def check(windows):
    if messagebox.askokcancel('提示', '要执行此操作吗') == 1:
        windows.destroy()  # 关闭窗口


def get_image(filename, width, height):
    im = Image.open(filename).resize((width, height))  # resize调整图片大小
    return ImageTk.PhotoImage(im)


# 表格样式1
def get_style_01(bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = "宋体"
    font.bold = bold
    font.underline = False
    font.italic = False
    font.colour_index = 0
    font.height = 300  # 200为10号字体
    style.font = font

    # 单元格居中
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    align.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
    style.alignment = align
    border = xlwt.Borders()  # 给单元格加框线
    border.left = xlwt.Borders.THIN  # 左
    border.top = xlwt.Borders.THIN  # 上
    border.right = xlwt.Borders.THIN  # 右
    border.bottom = xlwt.Borders.THIN  # 下
    style.borders = border
    return style


# 表格样式2
def get_style_02():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = "宋体"
    font.underline = False
    font.italic = False
    font.colour_index = 0
    font.height = 200  # 200为10号字体
    style.font = font

    # 单元格居中
    align = xlwt.Alignment()
    align.horz = xlwt.Alignment.HORZ_CENTER  # 水平方向
    align.vert = xlwt.Alignment.VERT_CENTER  # 竖直方向
    style.alignment = align
    border = xlwt.Borders()  # 给单元格加框线
    border.left = xlwt.Borders.THIN  # 左
    border.top = xlwt.Borders.THIN  # 上
    border.right = xlwt.Borders.THIN  # 右
    border.bottom = xlwt.Borders.THIN  # 下
    style.borders = border
    return style


def body():
    reader()
    createExcel()
    allocate()
    writer()
    messagebox.showinfo('提示', '完成分配')


def main():
    # 设置窗口，静止放大缩小，设置标题，设置窗体大小和在屏幕的位置
    windows = Tk()
    windows.title("考试监考系统")
    windows.resizable(False, False)
    windows.geometry('400x200+500+200')

    # 创建画布，将图片放在画布上
    canvas = Canvas(windows, width=400, height=200)
    im = get_image('logo.png', 400, 200)
    canvas.create_image(200, 100, image=im)
    # canvas.create_text(210, 55, fill='green', text='考试监考系统', font=('华文行楷', 30))
    canvas.pack()

    buttom1 = Button(canvas, text='开始分配', bg='lightgreen', fg='black', font=('黑体', 15),
                     command=lambda: body())  # lambda（）函数用于将特定数据发送到回调函数。
    buttom1.lift  # 将按钮上调到主界面不被Canvas覆盖
    buttom1.place(x=60, y=130)

    buttom2 = Button(canvas, text='查看结果', bg='lightgreen', fg='black', font=('黑体', 15), command=lambda: open_word())
    buttom2.lift
    buttom2.place(x=260, y=130)

    # protocol协议，WM_DELETE_WINDOW窗体关闭,这段代码是让窗体右上角x按钮点击关闭
    windows.protocol('WM_DELETE_WINDOW', lambda: check(windows))

    windows.mainloop()


def test():
    reader()


if __name__ == "__main__":
    # createExcel()
    # write_data_morning()
    # write_data_afternoon()
    # write_items()
    # write_classroom()
    # write_teachers()
    #
    # print("创建成功！")
    # createExcel()
    # reader()
    # createExcel()
    # print(teacherList)
    # print(subjectList)
    # print(classroomList)
    # bubbleSort(subjectList)
    # allocate()
    # writer()
    # arr = [1, 4, 7, 9, 34, 1, 4, 2, 5]
    #
    # print(quickSort(arr))
    main()
    # test()
