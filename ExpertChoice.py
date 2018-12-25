import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
# import pymysql
import pandas as pd

# from PIL import Image, ImageTk
class expertChoice:
    def __init__(self):

        # 获取数据
        # self.conn=pymysql.connect(host='localhost',port=3306,user='root',passwd='751982THUNDERlz',db='leizquant')
        # self.sql='select * from stocklist'
        # self.df=pd.read_sql(self.sql,self.conn)
        # self.conn.close()

        # 统一的对齐方式
        self.duiqi='W'
        self.root = Tk()
        self.root.title("中捷专家抽取系统")
        self.root.resizable(0, 0)
        self.root.geometry('1024x600')

        # 表头
        self.topimg=PhotoImage(file='top.gif')
        self.banner=ttk.Label(self.root,image=self.topimg)
        self.banner.grid(row=0,column=0)

        # 四个框架
        self.base = Frame(self.root,width=1000, height=200)
        self.show = Frame(self.root,width=900, height=200)
        self.manage = Frame(self.root, width=900, height=200)

        # 定义基本信息区域
        # 第一行
        Label(self.base,text='招标组织方式:').grid(row=0,column=0,sticky=self.duiqi)
        self.org = ttk.Combobox(self.base)
        self.org.grid(row=0, column=1,sticky=W)
        Label(self.base, text='招标人名称:').grid(row=0,column=2,sticky=self.duiqi)
        self.boss = Entry(self.base)
        self.boss.grid(row=0, column=3,sticky=self.duiqi)
        Label(self.base, text='招标代理机构项目编号:').grid(row=0, column=4,sticky=self.duiqi)
        self.agentid = Entry(self.base)
        self.agentid.grid(row=0, column=5,sticky=W)


        # 第二行
        Label(self.base, text='项目编号:').grid(row=1, column=0,sticky=self.duiqi)
        self.projectid = Entry(self.base,width=55)
        self.projectid.grid(row=1, column=1,columnspan=3,sticky=self.duiqi)
        Label(self.base, text='招标代理机构项目名称:').grid(row=1, column=4,sticky=self.duiqi)

        self.agentnames=StringVar()
        self.agentname_cb = ttk.Combobox(self.base,height=1,textvariable=self.agentnames)
        self.agentname_cb['values']=('中捷通信有限公司','公诚咨询管理有限公司')
        self.agentname_cb.grid(row=1, column=5,sticky=self.duiqi)

        # 第三行
        Label(self.base, text='项目名称:').grid(row=2, column=0,sticky=self.duiqi)
        self.projectname = Entry(self.base, width=55)
        self.projectname.grid(row=2, column=1, columnspan=3,sticky=self.duiqi)
        Label(self.base, text='递补专家人数:').grid(row=2, column=4,sticky=self.duiqi)

        self.agentnames = StringVar()
        self.agentname_cb = ttk.Combobox(self.base, height=1, textvariable=self.agentnames)
        self.agentname_cb['values'] = (1,2,3)
        self.agentname_cb.grid(row=2, column=5,sticky=self.duiqi)

        # 第四行
        Label(self.base, text='抽取地点:').grid(row=3, column=0,sticky=self.duiqi)
        self.choiceplace = Entry(self.base, width=55)
        self.choiceplace.grid(row=3, column=1, columnspan=3,sticky=self.duiqi)
        Label(self.base, text='抽取时间:').grid(row=3, column=4,sticky=self.duiqi)

        self.choicedate = StringVar()
        self.choicedate_cb = ttk.Combobox(self.base, height=1, textvariable=self.agentnames)
        self.choicedate_cb['values'] = (1, 2, 3)
        self.choicedate_cb.grid(row=3, column=5,sticky=self.duiqi)

        # 第五行
        Label(self.base, text='评标委员会人数:').grid(row=4, column=0,sticky=self.duiqi)
        self.confnum = StringVar()
        self.confnum_cb = ttk.Combobox(self.base, height=1, textvariable=self.confnum)
        self.confnum_cb['values'] = (1, 2, 3)
        self.confnum_cb.grid(row=4, column=1,sticky=self.duiqi)

        Label(self.base, text='业主代表人数:').grid(row=4, column=2,sticky=self.duiqi)
        self.bossnum = StringVar()
        self.bossnum_cb = ttk.Combobox(self.base, height=1, textvariable=self.bossnum)
        self.bossnum_cb['values'] = (1, 2, 3)
        self.bossnum_cb.grid(row=4, column=3,sticky=self.duiqi)

        self.bossinfo_btn=Button(self.base,text='填写业主代表信息')
        self.bossinfo_btn.grid(row=4, column=4,columnspan=2,rowspan=2,ipadx=10,ipady=10)

        # 第六行
        Label(self.base, text='专家评委人数:').grid(row=5, column=0,sticky=self.duiqi)
        self.expertnum = Entry(self.base)
        self.expertnum.grid(row=5, column=1,sticky=self.duiqi)

        Label(self.base, text='需抽取专家\n（含递补）人数:').grid(row=5, column=2,sticky=self.duiqi)
        self.expertchoicenum = Entry(self.base)
        self.expertchoicenum.grid(row=5, column=3,sticky=self.duiqi)


        # 展示抽取条件列表区域
        self.tree = ttk.Treeview(self.show, show="headings", height=5, columns=("a", "b"))
        self.vbar = ttk.Scrollbar(self.show, orient=VERTICAL, command=self.tree.yview)
        # 定义树形结构与滚动条
        self.tree.configure(yscrollcommand=self.vbar.set)

        # 表格的标题
        self.tree.column("a", width=200, anchor="center")
        self.tree.column("b", width=200, anchor="center")
        self.tree.heading("a", text="编号")
        self.tree.heading("b", text="名称")


        # 调用方法获取表格内容插入

        self.tree.grid(row=0, column=0, sticky=NSEW,rowspan=4,ipadx=20)
        self.vbar.grid(row=0, column=1, sticky=NS,rowspan=4)


        # 展示筛选条件
        Label(self.show, text='字段名称:').grid(row=0, column=2,padx=50)
        self.choicefield = StringVar()
        self.choicefield_cb = ttk.Combobox(self.show, height=1, textvariable=self.choicefield)

        self.choicefield_cb.grid(row=0, column=3)

        self.choicefield_cb['values'] = (1, 2, 3)


        Label(self.show, text='筛选条件1:').grid(row=1, column=2)
        self.choicecondition1 = StringVar()
        self.choicecondition1_cb = ttk.Combobox(self.show, height=1, textvariable=self.choicecondition1)
        # self.choicecondition1_cb['values'] = (1, 2, 3)
        self.choicecondition1_cb.grid(row=1, column=3)

        Label(self.show, text='筛选条件2').grid(row=2, column=2)
        self.choicecondition2 = StringVar()
        self.choicecondition2_cb = ttk.Combobox(self.show, height=1, textvariable=self.choicecondition2)
        # self.choicecondition2_cb['values'] = (1, 2, 3)
        self.choicecondition2_cb.grid(row=2, column=3)

        self.addcondition_btn=Button(self.show,text='添加筛选条件')
        self.addcondition_btn.grid(row=3,column=2)
        self.delcondition_btn = Button(self.show, text='删除筛选条件')
        self.delcondition_btn.grid(row=3, column=3)


        # 管理区域
        self.importdata_btn=Button(self.manage,text='导入数据',command=self.getdata)
        self.importdata_btn.grid(row=0,column=0,ipadx=30,padx=10,sticky='EW')

        self.choice_btn = Button(self.manage, text='开始抽取')
        self.choice_btn.grid(row=0, column=1,ipadx=30,padx=10,sticky='WE')

        self.check_btn = Button(self.manage, text='查看结果',command=self.test)
        self.check_btn.grid(row=0, column=2,ipadx=30,padx=10,sticky='WE')

        self.reset_btn = Button(self.manage, text='重置结果')
        self.reset_btn.grid(row=0, column=3,ipadx=30,padx=10,sticky='WE')

        Label(self.manage,text='目前已抽取人数：').grid(row=0,column=4,padx=30,sticky='WE')

        self.exit_btn = Button(self.manage, text='退出程序',command=exit)
        self.exit_btn.grid(row=0, column=5,ipadx=30,padx=10,sticky='WE')



        # 整体区域定位
        self.base.grid(row=1, column=0,pady=20,padx=12)
        self.show.grid(row=2, column=0)
        self.manage.grid(row=3, column=0)


        self.base.grid_propagate(0)
        self.show.grid_propagate(0)
        self.manage.grid_propagate(0)


        self.root.mainloop()
        # 表格内容插入
    def get_tree(self):

        for index,row in self.df.iterrows():
            self.tree.insert("", "end", values=(row[0],row[1]))
        # pass

    def __del__(self):
        pass

    def getdata(self):
        filename = filedialog.askopenfilename()
        self.df=pd.read_excel(filename,sheet_name='Sheet1',header=1,index_col=0)
        self.get_tree()
        self.choicefield_cb['values']=list(self.df['毕业院校'].unique())

    def test(self):
        print(self.choicefield.get())



if __name__ == '__main__':
     expertChoice()


