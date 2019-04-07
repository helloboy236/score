# score
用python编写的教务处查成绩的电脑窗体应用，并生成exe文件

Python窗体之大学生成绩查询
# 一、麻雀虽小，五脏俱全
本系统用到的python库：
tkinter——用于创建windows窗体
requests——用于请求获取数据
PIL——用于处理验证码图片
lxml——用于处理获得的数据
xlwt——用于将数据保存为Excel文件
pyinstaller——用于将python文件转为exe文件
# 二、思维导图
1）分析需求，建立大致方向
2）学习相关知识，学以致用
3）分析网页数据，并记录关键点
4）建立测试代码，得到满意结果为止
5）根据需求建立窗体大致模型
6）组合代码，不断运行
7）处理bug
8）生成exe
9）测试exe
10）分享exe
# 三、经验总结
## 1、窗体问题
### 1）创建窗体
from tkinter import *
root = Tk()
root.title("成绩查询系统")    #标题
root.geometry('360x300')
root.maxsize(360,300)       # 限制窗口大小
root.minsize(360,300)
root.mainloop()

简简单单五行代码就创建了一个不含任何控件的大小为360x300（中间为小写的x）的窗体


窗体属性
### 2）窗体中的控件
#定义退出按钮命令
def exit():
    root.quit()
btn1 = Button(root, text='退出系统', command=exit)
btn1.grid(column=0,row=0)
定义了一个按钮，放在root窗体，显示名称为 退出系统， command理解为按钮点击时执行的函数；.grid()为一种排列方式,另外还有  .place()  ,  .pack()

窗体控件


### 3）动态设置控件属性(文本)
zh = Entry(root)
zh.grid(column=3, row=2)
zh.insert(0,'2016200000') #通过该语句可向输入框中插入字符串


keyword = Entry(root,show='*') # 若为密码则输入的字符将显示为 *
keyword.grid(column=3, row=4)

keyword['show']=''  # 该语句会修改显示方式(即显示输入的文本)

### 4）制作表格
from tkinter import ttk, messagebox
from tkinter import *
root = Tk()
root.title('全部成绩')
root.geometry('800x600')
root.maxsize(800,600)
root.minsize(800,600)

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

root.mainloop()


## 2、爬虫问题
import requests
# 请求头
header = {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:66.0) Gecko/20100101 Firefox/66.0'
    }
# 请求数据
data = {
    'zjh':'',
    'mm':'',
    'v_yzm':''
}

brower = requests.session() # 创建session对象

cookie = brower.cookies.get_dict()  # 获取cookie值

res = brower.post(url1,data=data,headers=header,cookies=cookie).text 



## 3、保存到Excel
import xlwt

columns = ('学期','科目','学分','课程属性','成绩')

dir = '我的成绩单.xlsx'   # 文件路径
f = xlwt.Workbook()    # 创建工作簿
sheet1 = f.add_sheet('成绩单') #创建成绩单工作表
sheet1.col(0).width = 256*50 #设置单元格宽度
sheet1.col(1).width = 256 * 50
sheet1.col(2).width = 256 * 10
sheet1.col(3).width = 256 * 10
sheet1.col(4).width = 256 * 10
row0 = list(columns) 
for i in range(5):
    sheet1.write(0,i,row0[i]）# 向工作表第一行写数据作为表头
f.save(dir)  # 保存文件到指定路径

## 4、生成exe
### 1）安装pyinstaller

pip install pyinstaller

### 2）打包
cmd 命令行切换到项目目录，执行下列语句

pyinstaller -F -w main.py   # 此处应注意大小写

最后一行看到successful字样，即可在项目根目录的dist目录下找到exe文件，双击即可执行
### 3）部分错误
若exe文件因错误无法打开
尝试以下语句
pyinstaller -F main.py   # 此处应注意大小写
完成后双击exe文件会有一个黑框一闪而过，需要用到截屏将黑框截图，里面显示了错误信息，根据错误信息对症下药。


参数说明
## 四、全部代码
### 1、代码已上传到GitHub
  https://github.com/helloboy236/score
### 2、pyinstaller 生成的exe文件包
链接：https://pan.baidu.com/s/1LClBmu54FlwvZcOvgbk2-A 
提取码：izp3 
复制这段内容后打开百度网盘手机App，操作更方便哦
