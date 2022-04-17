"""
 -*- coding: utf-8 -*-

 @Time : 2021/11/3 8:51

 @Author : jagger

 @File : 考试监考系统.py

 @Software: PyCharm 

 @contact: 252587809@qq.com

 -*- 功能说明 -*-
文件读写功能
考场分配功能
可视化界面
"""
import xlrd
import xlwt
import random


class struct_teacher:
    def __init__(self):
        self.id = ''
        self.dept = ''
        self.name = ''
        self.sex = ''


class struct_subject:
    def __init__(self):
        self.id = ''
        self.subject = ''
        self.amount = ''


class struct_classroom:
    def __init__(self):
        self.id = ''
        self.classroom = ''
        self.amount = ''


class allocate:
    def __init__(self):
        self.source_file_name = '监考人员-科目-教师名单.xls'
        self.excelName = '监考分配表'
        # 每个列表的第0行存的是提示信息
        # 内容是结构体对象
        self.teacherList = []  # 监考人员名单
        self.subjectList = []  # 考试科目
        self.classroomList = []  # 空闲教室

    def reader(self):
        '''
        读取信息文件
        Returns:

        '''
        work_book = xlrd.open_workbook(self.source_file_name)
        # for items in work_book.sheet_names():
        #     if items == '监考人员名单':
        #         # for item in items:
        #         print(items.row.values())
        #
        #     elif items == '考试科目':
        #         print('2')
        #     elif items == '空闲教室':
        #         print('3')
        #     else:
        #         print('wrong!')
        # tableList = ['监考人员名单', '考试科目', '空闲教室']
        teacher_table = work_book.sheet_by_name('监考人员名单')
        subject_table = work_book.sheet_by_name('考试科目')
        classroom_table = work_book.sheet_by_name('空闲教室')
        # for tableName in tableList:
        #     table = work_book.sheet_by_name(tableName)
        #     for row in table:
        #         if tableName == '监考人员名单':
        #             teacher = struct_teacher()
        #             teacher.id = row[0].value
        #             teacher.dept = row[1].value
        #             teacher.sex = row[2].value
        #             teacher.name = row[3].value
        #             self.teacherList.append(teacher)
        for i in teacher_table:
            teacher = struct_teacher()
            # subject = struct_subject()
            # classroom = struct_classroom()
            # teacher.id = i.row.value(1, 1)
            # i.row.value(1,1)
            # teacher.id = i
            # teacher.dept = j
            # teacher.sex = k
            # teacher.name = l
            # print(i[0].value)
            #     self.teacherList.append(teacher)
            # print(self.teacherList)
            teacher.id = i[0].value
            teacher.dept = i[1].value
            teacher.sex = i[2].value
            teacher.name = i[3].value
            self.teacherList.append(teacher)
        for i in subject_table:
            subject = struct_subject()
            subject.id = i[0].value
            subject.subject = i[1].value
            subject.amount = i[2].value
            self.subjectList.append(subject)
        for i in classroom_table:
            classroom = struct_classroom()
            classroom.id = i[0].value
            classroom.classroom = i[1].value
            classroom.amount = i[2].value
            self.subjectList.append(classroom)

    def writer(self):
        '''
        将分配结果写入文件
        Returns:

        '''
        work_book = xlwt.Workbook(encoding='utf-8')  # 创建工作簿
        j = 1
        for i in self.classList:

            table = work_book.add_sheet('%s 班' % j)  # 添加工作表
            j += 1
            # 向工作表写入数据
            table.write(0, 0, '姓名:')  # 写第零行的属性
            table.write(0, 1, '性别:')
            k = 1
            for stu in i:
                table.write(k, 0, stu['name'])
                table.write(k, 1, stu['sex'])
                # table.write(k, 2, stu['flag'])
                k += 1

        work_book.save(self.excelName)  # 保存工作簿

    # def produce_stuList(self):
    #     '''
    #     生成学生信息
    #     数据结构为【{}】列表套字典 即用列表储存学生 字典中存放的是学生的信息：名字、性别、是否分了班
    #     Returns:
    #
    #     '''
    #     for i in range(1200):
    #         _dict = {}  # 列表套字典
    #         _dict['name'] = random.choice(self.firstNameList) + random.choice(self.secondNameList)
    #         _dict['sex'] = random.choice(['男', '女'])
    #         # _dict['flag'] = False  # 是否完成分班
    #         self.stuList.append(_dict)

    def allocation(self):
        '''
        实现监考考场分配
        将学生分配到班级中去
        数据结构为【【{}】】 列表套列表再套字典 最外层列表放的是班级共二十二个。第二层列表放的是学生，每班大约1200/22人。字典存放的是学生信息
        classList[_class[stu{}]]
        Returns:

        '''
        for i in range(22):
            _class = []
            stuList = self.stuList
            for j in range(1200 // 22):
                stu = random.choice(stuList)
                _class.append(stu)
                stuList.remove(stu)
            self.classList.append(_class)
        for i in stuList:
            __class = random.choice(self.classList)
            __class.append(i)


def main():
    mytool = tools()
    mytool.reader()


if __name__ == '__main__':
    main()
