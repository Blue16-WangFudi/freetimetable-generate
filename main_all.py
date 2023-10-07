#encoding=utf-8
#author: blue16（blue16.cn）
#date: 2023-10-5
#summary: 遍历学生会所有部门的课程表（列表样式）并自动生成无课表

import tabula
import csv
import openpyxl
import os

#将一个pdf转为csv
def convert_to_csv(pdf_path, csv_path):
    tabula.convert_into(pdf_path, csv_path, output_format="csv", pages="all")

#读取指定的列,并清除无效数据
def read_column(csv_path,n):#csv_path=example.csv
    with open(csv_path) as csvfile:
        reader = csv.reader(csvfile)
        templist=[]
        for row in reader:
            temp=row[n]
            #print(temp,len(temp))
            if len(temp)>=3 and len(temp)<=5:#加一个判断，剔除无效数据
                templist.append(temp) # 提取第n列数据read_column(1)
        return templist

#写一条记录
def write_record(wb_sheet,department,name,week_num,section_num):
    #num均从1开始，方便观察;索引从0开始递增
    week_list=['C','D','E','F','G','H','I']
    section_list=list(range(2,16))#左开右闭一定注意
    cell_id=week_list[week_num-1]+str(section_list[section_num-1])
    if wb_sheet[cell_id].value==None:
        wb_sheet[cell_id]=department+' '+name+'，'
    else:
        wb_sheet[cell_id]=wb_sheet[cell_id].value+department+' '+name+'，'



#解析
#想法：用一个list，每一天从1-14，14个cell，有课为1，无课为0，然后遍历写入
#先将一个pair转为int的list，然后根据大小变化即可划分第几周
def pair_to_numberlist(pair_list):#把类似于"A-B"的pair转为[A,B]的list
    templist=[]
    for s in pair_list:#逐个取出课程时间范围并打印
        templist.append(s.split("-"))
    return templist


def parse_numberlist(numberlist):#解析并将上课情况写入一个二维数组
    pre=0
    #外层：周几；内层:1-14
    week=1
    prev_num=['0','0']
    #course_list=[[0]*14]*7#注意不能这么写，浅拷贝会导致改一个就把全部改了
    course_list=[[0 for i in range(14)] for j in range(7)]
    for num in numberlist:
        if(int(num[1])<int(prev_num[1])):
            week=week+1
        #print(num[0],num[1])
        for i in range(int(num[0]),int(num[1])+1):#左闭右开
            #print(week,'@',int(i))
            course_list[week-1][int(i)-1]=1#有数据证明有课
        prev_num=num
    return course_list

def output_a_member(wb_sheet,course_list,department,name):
    #写入一个成员的课表信息
    week=0
    section=0
    for temp_week in course_list:
        week=week+1
        for temp_section in temp_week:
            if course_list[week-1][section-1]!=1:
                write_record(wb_sheet,department,name,week,section)
            section=section+1
            #print(department,'#',name,week,section) 
        section=0
def welcome():#准备工作
    print('-----------------------------')
    print('无课表自动生成系统V1.0，西南大学人工智能学院团委办公室部门制作，作者Blue16，个人主页blue16.cn')
    input('请将所有数据按照格式要求放入InputTable文件夹中，程序解析结束可以指定输出位置。按回车继续')
    print('开始处理')
    print('-----------------------------')


def search_department_member(path,department):#搜索对应部门成员然后直接写入
    member_file_list=os.listdir(path+'\\InputTable\\'+department)
    member_list=[]
    for temp in member_file_list:
        temp2=temp.split('.')
        if(temp2[1]=='pdf'):
            member_list.append(temp)
    print('查找到的成员文件:',member_list)
    for temp in member_list:
        temp2=temp.split('(')
        print('当前写入部门：',department,"当前写入人员：",temp2[0])
        #准备工作
        sourcepath=path+'\\InputTable\\'+department+'\\'+temp
        destpath=path+'\\InputTable\\'+department+'\\'+temp2[0]+'.csv'
        convert_to_csv(sourcepath,destpath)
        mylist=read_column(destpath,1)
        numberlist=pair_to_numberlist(mylist)
        output_a_member(wb_sheet,parse_numberlist(numberlist),department,temp2[0])
    
welcome()
print('加载输出模板')
wb=openpyxl.load_workbook('output.xlsx')
wb_sheet=wb['Sheet']
path=os.getcwd()#D:\桌面\无课表
department_list=os.listdir(path+'\\InputTable')
print('查找到的部门：',department_list)
for temp in department_list:
    search_department_member(path,temp)
#output_a_member(wb_sheet,parse_numberlist(numberlist),'办公室','我')
outputpath=input('解析完成，请输入要保存的文件名,然后回车：（可以是路径）')
while outputpath=='':
    outputpath=input('文件名无效，请输入文件名后回车：（可以是路径）')
wb.save(outputpath)#注意保存
print('文件已经保存到'+outputpath+'。完成，撒花')
