
import xlrd
import xlwt

class Student:

    def __init__(self):
        self.id = ''
        self.name = ''
        self.type = ''
        self.belonging = ''
        self.profession = ''
        self.total_mark = 0
        self.exam_type = ''
        self.marks = []
        self.mark_type = []
        self.knowledge_point = {}

    def read_mark_type(self, worksheet):
        for i in range(1, worksheet.nrows):
            self.mark_type.append((int(worksheet.cell_value(i, 2)), worksheet.cell_value(i, 3)))
            #print(self.mark_type)
    def deal_knowledge_point(self):
        for i in range(1, 12):
            self.knowledge_point['知识点%d' % i] = [0, 0]
        print(self.knowledge_point)
        for i in range(len(self.marks)):
            self.knowledge_point[self.mark_type[i][1]][0] += self.marks[i]
            self.knowledge_point[self.mark_type[i][1]][1] += self.mark_type[i][0]
            #print(self.knowledge_point)

def output_knowledge_point(students, work_book):
    sheet1 = work_book.add_sheet('输出表')
    sheet1.write(0, 0, '序号')
    sheet1.write(0, 1, '姓名')
    sheet1.write(0, 2, '人员分类')
    sheet1.write(0, 3, '单位')
    sheet1.write(0, 4, '专业')
    sheet1.write(0, 5, '成绩')
    sheet1.write(0, 6, '试卷类型')
    for i in range(1, 12):
        sheet1.write(0, i + 6, '知识点%d' % i)
    for i in range(len(students)):
        sheet1.write(i+1, 0, students[i].id)
        sheet1.write(i+1, 1, students[i].name)
        sheet1.write(i+1, 2, students[i].type)
        sheet1.write(i+1, 3, students[i].belonging)
        sheet1.write(i+1, 4, students[i].profession)
        sheet1.write(i+1, 5, students[i].total_mark)
        sheet1.write(i+1, 6, students[i].exam_type)
        for j in range(1, 12):

            if (students[i].knowledge_point['知识点%d' % j][1]) == 0:
                sheet1.write(i + 1, j + 6, 0)
            else:
                # sheet1.write(i + 1, j + 6, '%d / %d' % (students[i].knowledge_point['知识点%d' % j][0], students[i].knowledge_point['知识点%d' % j][1]))
                 sheet1.write(i+1, j+6, '%.4f' %((students[i].knowledge_point['知识点%d' % j][0])/(students[i].knowledge_point['知识点%d' % j][1])))

def output_belonging_knowledge_point(students, work_book):
    sheet1 = work_book.add_sheet('单位表')
    renshu = {}
    
    mat = {}
    for student in students:
        if student.belonging in renshu:
            renshu[student.belonging] += 1
            # print(renshu)
        else:
            renshu[student.belonging] = 1
            mat[student.belonging] = {}
            for i in range(1, 12):
                mat[student.belonging]['知识点%d' % i] = 0
    for student in students:
        for i in range(1, 12):
            mat[student.belonging]['知识点%d' % i] += student.knowledge_point['知识点%d' % i][0]
    lst = []
    for item in renshu:
        lst.append(item)
    lst.sort()
    sheet1.write(0, 0, '单位')
    for i in range(1, 12):
        sheet1.write(0, i, '知识点%d' % i)
    for i in range(len(lst)):
        sheet1.write(i+1, 0, lst[i])
        for j in range(1, 12):
            sheet1.write(i+1, j, mat[lst[i]]['知识点%d' % j] / renshu[lst[i]])


def output_type_knowledge_point(students, work_book):
    sheet1 = work_book.add_sheet('人员类别')
    renshu = {}
    mat = {}
    for student in students:
        if student.type in renshu:
            renshu[student.type] += 1
        else:
            renshu[student.type] = 1
            mat[student.type] = {}
            for i in range(1, 12):
                mat[student.type]['知识点%d' % i] = 0
    for student in students:
        for i in range(1, 12):
            mat[student.type]['知识点%d' % i] += student.knowledge_point['知识点%d' % i][0]
    lst = []
    for item in renshu:
        lst.append(item)
    lst.sort()
    sheet1.write(0, 0, '人员类别')
    for i in range(1, 12):
        sheet1.write(0, i, '知识点%d' % i)
    for i in range(len(lst)):
        sheet1.write(i+1, 0, lst[i])
        for j in range(1, 12):
            sheet1.write(i+1, j, mat[lst[i]]['知识点%d' % j] / renshu[lst[i]])

def output_belonging_type_knowledge_point(students, work_book):
   sheet1 = work_book.add_sheet('人员类别+单位')
   renshu = {}
   mat = {}
   for student in students:
       if student.type in renshu:
           renshu[student.type] += 1
       else:
           renshu[student.type] = 1
           mat[student.type] = {}
           for i in range(1, 12):
               mat[student.type]['知识点%d' % i] = 0

def output_students(students):
    work_book = xlwt.Workbook('2018_out_sun.xlsx')
    output_knowledge_point(students, work_book)
    output_belonging_knowledge_point(students, work_book)
    output_type_knowledge_point(students, work_book)
    work_book.save('2018_out_sun.xls')

def read_students(students):
    #打开一个workbook
    workbook = xlrd.open_workbook('2018.xlsx')
    #定位到sheet1
    worksheet1 = workbook.sheet_by_name(u'成绩表')
    for i in range(1, worksheet1.nrows):
        student = Student()
        student.id = int(worksheet1.cell_value(i, 0))
        student.name = worksheet1.cell_value(i, 1)
        student.type = worksheet1.cell_value(i, 2)
        student.belonging = worksheet1.cell_value(i, 3)
        student.profession = worksheet1.cell_value(i, 4)
        student.total_mark = float(worksheet1.cell_value(i, 5))
        student.exam_type = worksheet1.cell_value(i, 6)
        for j in range(7, worksheet1.ncols):
            student.marks.append(int(worksheet1.cell_value(i, j)))
        student.read_mark_type(workbook.sheet_by_name('%s卷' % student.exam_type))
        student.deal_knowledge_point()
        students.append(student)
        print(student.marks)


if __name__ == '__main__':
    students = []
    read_students(students)
    output_students(students)



