from tkinter import ttk, messagebox
from tkinter import *
import xlwt
from tkinter.filedialog import askdirectory
from lxml import etree

# 整理html数据并调用自定义创建窗体函数进行显示
def show(html):
    z = 0  # 用于行计数
    htm = etree.HTML(html)
    tbodys = htm.xpath('//*[@id="user"]')
    #  //*[@id="user"]/tbody/tr[1]/td[1]
    titles = htm.xpath('//a/@name')
    for title in titles:
        pass
    subjects = []
    nums = len(tbodys)
    n = 0
    for n in range(nums):  # 全部学期
        subjects.append(titles[n])
        for i in tbodys[n].findall('tr'):  # 一学期
            subject = {}
            for j in i.findall('td'):  # 科目
                detail = j.text.strip()
                z += 1
                if (z == 3):
                    subject['course'] = detail
                elif (z == 5):
                    subject['credit'] = detail
                elif (z == 6):
                    subject['attribute'] = detail
                elif (z == 7):
                    subject['score'] = j.findall('p')[0].text.strip()
                else:
                    continue
            subjects.append(subject)
            z = 0
        n += 1
    createWin(subjects)

# 创建显示成绩单的窗体
def createWin(sub):
    root = Tk()
    root.title('全部成绩')
    root.geometry('800x600')
    root.maxsize(800,600)
    root.minsize(800,600)
    # 设置保存文件的单元格格式
    def set_style(name, height, bold=False):
        style = xlwt.XFStyle()  # 初始化样式
        font = xlwt.Font()  # 为样式创建字体
        font.name = name
        font.bold = bold
        font.color_index = 4
        font.height = height
        style.font = font
        al = xlwt.Alignment()
        al.horz = 0x02  # 设置水平居中
        al.vert = 0x01  # 设置垂直居中
        style.alignment = al
        return style
    # 当点击保存按钮时执行此函数
    def save():
        if(inp.get()=='桌面'):
            s = r"C:\Users\Administrator\Desktop"
        else:
            s = inp.get()
        dir = s+'\\我的成绩单.xlsx' # 合成文件路径
        f = xlwt.Workbook() # 创建工作簿
        sheet1 = f.add_sheet('成绩单') # 新建工作表，名为成绩单
        sheet1.col(0).width = 256*50    # 设置单元格的宽度
        sheet1.col(1).width = 256 * 50
        sheet1.col(2).width = 256 * 10
        sheet1.col(3).width = 256 * 10
        sheet1.col(4).width = 256 * 10
        row0 = list(columns)    # 表头数据
        for i in range(5):
            sheet1.write(0,i,row0[i],set_style('Times New Roman',220,True)) # 写第一行数据作为表头

        z = 1   # 行
        course = []
        for i in sub:
            if (type(i) == dict):
                if (course == []):
                    course.append('')
                for j in i:
                    course.append(i[j])
            else:
                course.append(i)
            if (type(i) == dict):
                for i in range(len(course)):
                    sheet1.write(z, i, course[i],set_style('宋体',220,False))
                course.clear()
            z += 1
        f.save(dir) # 保存文件到选定路径
        messagebox.showinfo('提示', '保存成功，请到保存目录查看《我的成绩单》')
    # 用于让用户自己选择成绩单保存地址
    def select():
        global s
        s = askdirectory()
        inp.delete(0,END)
        inp.insert(0,s)
    lb = Label(root,text='选择保存路径：')
    lb.place(x=0,y=0)
    inp = Entry(root,width=30)
    inp.place(x=100,y=0)
    inp.insert(0,'桌面')
    savefile = Button(root, text='重新选择保存路径', command=select)
    savefile.pack()
    savefile = Button(root,text='保存到本地',command=save)
    savefile.place(x=600,y=0)
    la = Label(root,pady=3)
    la.pack()
    global columns
    columns = ('学期','科目','学分','课程属性','成绩')

    tv = ttk.Treeview(root, show='headings',columns=columns)
    tv.column('学期', width=200, anchor='center')
    tv.column('科目',width=300,anchor='center')
    tv.column('学分',width=100,anchor='center')
    tv.column('课程属性',width=100,anchor='center')
    tv.column('成绩',width=100,anchor='center')

    tv.heading('学期', text='学期')
    tv.heading('科目',text='科目')
    tv.heading('学分',text='学分')
    tv.heading('课程属性',text='课程属性')
    tv.heading('成绩',text='成绩')
    tv.pack(side=LEFT,fill=BOTH)

    z=1
    course = []
    for i in sub:
        if(type(i)==dict):
            if(course==[]):
                course.append('')
            for j in i:
                course.append(i[j])
        else:
            course.append(i)
        cs = tuple(course)
        z += 1
        if(type(i)==dict):
            tv.insert('', z, values=cs)
            course.clear()
    root.mainloop()