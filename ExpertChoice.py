import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import sqlite3
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
        self.root.geometry('960x600')

        # 全局变量
        self.selectedexpert=0

        # 初始化数据库
        self.dbconn=sqlite3.connect('mydb.db')
        self.dbcur=self.dbconn.cursor()
        # self.dbcur.execute('create table if not exists tbexpert(id)')
        # self.dbcur.execute('create table if not exists tbcata(id)')

        # 表头
        self.topimg=PhotoImage(file='top.gif')
        self.banner=ttk.Label(self.root,image=self.topimg)
        self.banner.grid(row=0,column=0,sticky='NWSE')

        # 四个框架
        self.base = Frame(self.root,width=960, height=200)
        self.show = Frame(self.root,width=960, height=120)
        self.rlt = Frame(self.root, width=960, height=130)
        self.manage = Frame(self.root, width=960, height=100)

        # 定义基本信息区域
        # 第一行
        Label(self.base,text='招标组织方式:').grid(row=0,column=0,sticky=self.duiqi)
        self.org_var=StringVar()
        self.org_cb = ttk.Combobox(self.base,textvariable=self.org_var)
        self.org_cb['values']=['公开招标','邀请招标','公开询价','单一来源采购']
        self.org_cb.current(0)
        self.org_cb.grid(row=0, column=1,sticky=W)


        Label(self.base, text='招标人名称:').grid(row=0,column=2,sticky=self.duiqi)
        self.boss = Entry(self.base)
        self.boss.grid(row=0, column=3,sticky=self.duiqi)
        Label(self.base, text='招标代理机构\n项目编号:').grid(row=0, column=4,sticky=self.duiqi)
        self.agentid = Entry(self.base)
        self.agentid.grid(row=0, column=5,sticky=W)


        # 第二行
        Label(self.base, text='项目编号:').grid(row=1, column=0,sticky=self.duiqi)
        self.projectid = Entry(self.base,width=55)
        self.projectid.grid(row=1, column=1,columnspan=3,sticky=self.duiqi)

        Label(self.base, text='招标代理机构\n项目名称:').grid(row=1, column=4,sticky=self.duiqi)
        self.agentnames=StringVar()
        self.agentname_cb = ttk.Combobox(self.base,height=1,textvariable=self.agentnames)
        self.agentname_cb['values']=('中捷通信有限公司','公诚咨询管理有限公司')
        self.agentname_cb.current(0)
        self.agentname_cb.grid(row=1, column=5,sticky=self.duiqi)

        # 第三行
        Label(self.base, text='项目名称:').grid(row=2, column=0,sticky=self.duiqi)
        self.projectname = Entry(self.base, width=55)
        self.projectname.grid(row=2, column=1, columnspan=3,sticky=self.duiqi)


        Label(self.base, text='递补专家人数:').grid(row=2, column=4,sticky=self.duiqi)
        self.backexpertnum_var = StringVar()
        self.backexpertnum_cb = ttk.Combobox(self.base, height=1, state='readonly',textvariable=self.backexpertnum_var)
        self.backexpertnum_cb['values'] = (0,1,2,3)
        self.backexpertnum_cb.current(0)
        self.backexpertnum_cb.grid(row=2, column=5,sticky=self.duiqi)

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
        self.confnum_var = StringVar()
        self.confnum_cb = ttk.Combobox(self.base,textvariable=self.confnum_var)
        self.confnum_cb['values']=(5,7,9,11,13,15,17,19)
        self.confnum_cb.current(0)
        self.confnum_cb.grid(row=4, column=1,sticky=self.duiqi)
        # self.confnum_et.bind('<FocusOut>', self.checkconfnum())

        Label(self.base, text='业主代表人数:').grid(row=4, column=2, sticky=self.duiqi)
        self.bossnum_var = StringVar()
        self.bossnum_cb = ttk.Combobox(self.base,state='readonly', height=1, textvariable=self.bossnum_var)
        self.bossnum_cb['values'] = (1, 0)
        self.bossnum_cb.current(1)
        self.bossnum_cb.grid(row=4, column=3,sticky=self.duiqi)

        self.bossinfo_btn=Button(self.base,text='填写业主代表信息')
        self.bossinfo_btn.grid(row=4, column=4,columnspan=2,rowspan=2,ipadx=10,ipady=10)

        # 第六行
        Label(self.base, text='专家评委人数:').grid(row=5, column=0,sticky=self.duiqi)
        self.expertnum_var = StringVar()
        self.expertnum = Entry(self.base,state='readonly',bg='#808080',textvariable=self.expertnum_var)
        self.expertnum.grid(row=5, column=1,sticky=self.duiqi)

        Label(self.base, text='需抽取专家\n（含递补）人数:').grid(row=5, column=2,sticky=self.duiqi)
        self.expertchoicenum_var = StringVar()
        self.expertchoicenum = Entry(self.base,state='readonly',textvariable=self.expertchoicenum_var)
        self.expertchoicenum.grid(row=5, column=3,sticky=self.duiqi)


        # 展示抽取条件列表区域
        self.tree = ttk.Treeview(self.show, show="headings", height=5, columns=("a", "b"))
        self.vbar = ttk.Scrollbar(self.show, orient=VERTICAL, command=self.tree.yview)
        # 定义树形结构与滚动条
        self.tree.configure(yscrollcommand=self.vbar.set)
        # 表格的标题
        self.tree.column("a", width=150, anchor="center")
        self.tree.column("b", width=150, anchor="center")
        self.tree.heading("a", text="筛选字段")
        self.tree.heading("b", text="筛选条件")
        # 调用方法获取表格内容插入
        self.tree.grid(row=0, column=0, sticky=NSEW,rowspan=4)
        self.vbar.grid(row=0, column=1, sticky=NS,rowspan=4)


        # 展示筛选条件
        Label(self.show, text='字段名称:').grid(row=0, column=2,padx=50,sticky=self.duiqi)
        self.choicefield = StringVar()
        self.choicefield_cb = ttk.Combobox(self.show, height=1, textvariable=self.choicefield)
        self.choicefield_cb.grid(row=0, column=3,sticky=self.duiqi)
        self.choicefield_cb['values'] = (1, 2, 3)


        # Label(self.show, text='筛选条件1:').grid(row=1, column=2)
        self.choicecondition1 = StringVar()
        self.choicecondition1_cb = ttk.Combobox(self.show, height=1, textvariable=self.choicecondition1)
        # self.choicecondition1_cb['values'] = (1, 2, 3)
        # self.choicecondition1_cb.grid(row=1, column=3,sticky=self.duiqi)

        # Label(self.show, text='筛选条件2').grid(row=2, column=2)
        self.choicecondition2 = StringVar()
        self.choicecondition2_cb = ttk.Combobox(self.show, height=1, textvariable=self.choicecondition2)
        # self.choicecondition2_cb['values'] = (1, 2, 3)
        # self.choicecondition2_cb.grid(row=2, column=3,sticky=self.duiqi)

        self.addcondition_btn=Button(self.show,text='添加筛选条件',width=15,height=1)
        self.addcondition_btn.grid(row=3,column=2)
        self.delcondition_btn = Button(self.show, text='删除筛选条件',width=15,height=1)
        self.delcondition_btn.grid(row=3, column=3)

        self.choicecondition_lb=Listbox(self.show,height=6)
        self.choicecondition_lb.grid(row=0,column=4,rowspan=4,padx=10,sticky=self.duiqi)

        # 抽取结果
        Label(self.rlt, text='目前已抽取人数：',width=68,anchor='w').grid(row=0, column=0, sticky=self.duiqi)

        self.export=Button(self.rlt,text='导出结果',width=15)
        self.export.grid(row=0,column=2,sticky='E')

        self.reset_btn = Button(self.rlt, text='重置结果',width=15)
        self.reset_btn.grid(row=0, column=1,sticky='E')


        self.rlttree = ttk.Treeview(self.rlt, show="headings", height=5, columns=("a", "b", "c", "d", "e", "f"))
        self.rltvbar = ttk.Scrollbar(self.rlt, orient=VERTICAL, command=self.rlttree.yview)
        # 定义树形结构与滚动条
        self.rlttree.configure(yscrollcommand=self.vbar.set)
        # 表格的标题
        self.rlttree.column("a", width=80, anchor="center")
        self.rlttree.column("b", width=50, anchor="center")
        self.rlttree.column("c", width=200, anchor="center")
        self.rlttree.column("d", width=200, anchor="center")
        self.rlttree.column("e", width=180, anchor="center")
        self.rlttree.column("f", width=200, anchor="center")
        self.rlttree.heading("a", text="姓名")
        self.rlttree.heading("b", text="性别")
        self.rlttree.heading("c", text="现任职务")
        self.rlttree.heading("d", text="工作单位")
        self.rlttree.heading("e", text="电话")
        self.rlttree.heading("f", text="身份证号")
        # 调用方法获取表格内容插入
        self.rlttree.grid(row=1, column=0, sticky='NEW',columnspan=3)
        self.rltvbar.grid(row=1, column=3, sticky='NS',columnspan=3)


        # 管理区域
        self.importdata_btn=Button(self.manage,text='导入专家库',command=self.importdata_func)
        self.importdata_btn.grid(row=0,column=0,ipadx=30,padx=0,pady=5,sticky='W')

        self.cleardata_btn = Button(self.manage, text='清空专家库',command=self.cleardata_func)
        self.cleardata_btn.grid(row=0, column=1,ipadx=30,padx=0,sticky='W')

        self.checkdata_btn = Button(self.manage, text='查看专家库',command=self.showexpert_func)
        self.checkdata_btn.grid(row=0, column=2,ipadx=30,padx=0,sticky='W')



        self.exit_btn = Button(self.manage, text='退出程序',width=30,command=exit)
        self.exit_btn.grid(row=0, column=3,ipadx=30,padx=120,sticky='E')



        # 整体区域定位
        self.base.grid(row=1, column=0,pady=10,padx=12)
        self.show.grid(row=2, column=0)
        self.rlt.grid(row=3, column=0)
        self.manage.grid(row=4, column=0,pady=10)


        self.base.grid_propagate(0)
        self.show.grid_propagate(0)
        self.rlt.grid_propagate(0)
        self.manage.grid_propagate(0)


        self.showconfnum_func()
        self.bindshowconfnum_func()
        self.root.mainloop()

    # 绑定各种数字的刷新
    def bindshowconfnum_func(self):
        self.bossnum_cb.bind('<<ComboboxSelected>>',self.showconfnum_func)
        self.confnum_cb.bind('<<ComboboxSelected>>', self.showconfnum_func)
        self.backexpertnum_cb.bind('<<ComboboxSelected>>', self.showconfnum_func)

        # 表格内容插入
    def showconditiontree_func(self):
        # dfcondition
        # for index,row in self.df.iterrows():
        #     self.tree.insert("", "end", values=(row[0],row[1]))
        pass

    def showrlttree_func(self):
        # for index,row in self.df.iterrows():
        #     self.tree.insert("", "end", values=(row[0],row[1]))
        pass

    # 清空树
    def cleartree_func(tree):
        x = tree.get_children()
        for item in x:
            tree.delete(item)


    def __del__(self):
        self.dbconn.close()


    # 导入专家库
    def importdata_func(self):
        filename = filedialog.askopenfilename()
        if filename!='':
            try:
                xlsfile=pd.ExcelFile(filename)
                self.dfexpert=xlsfile.parse(sheet_name='专家名单',header=0,index_col=0)
                self.dfexpert = self.dfexpert.astype(str)
                self.dfexpert.to_sql('tbexpert',self.dbconn,if_exists='replace')

                self.dfcata = xlsfile.parse(sheet_name='专业分类', header=0)
                self.dfcata['编号'].fillna(method='ffill', inplace=True)
                self.dfcata.set_index('编号',inplace=True)
                self.dfcata['专业分类'].fillna(method='ffill', inplace=True)
                self.dfcata.to_sql('tbcata', self.dbconn, if_exists='replace')

                messagebox.showinfo(title='导入成功', message='导入成功，可以按查看专家按钮查看。！')
                # self.showconditiontree_func()
                # self.choicefield_cb['values']=list(self.dfcata['专业分类'].unique())
            except:
                messagebox.showwarning(title='导入失败',message='导入失败，请重新导入！')


    # 清空专家库
    def cleardata_func(self):
        self.dfnull=pd.DataFrame()
        self.dfcata=self.dfnull
        self.dfexpert=self.dfnull
        self.dfnull.to_sql('tbexpert',self.dbconn,if_exists='replace')
        self.dfnull.to_sql('tbcata', self.dbconn, if_exists='replace')
        messagebox.showinfo(title='清空专家库',message='专家库已经清空！')


    def test(self):
        print(self.choicefield.get())

    def checknum_func(self):
        try:
            int(self.confnum_var.get())
        except:
            messagebox.showwarning(title='注意评标委员为人数',message='请填入3以上单数')
            return False
        if self.confnum_var.get()=='':
            messagebox.showwarning(title='注意评标委员为人数',message='不能为空')
            return False
        if int(self.confnum_var.get())<=3 or int(self.confnum_var.get())%2==0:
            messagebox.showwarning(title='注意评标委员为人数', message='不符合3人以上单数的要求')
            return False
        if self.bossnum_var.get()=='' :
            messagebox.showwarning(title='注意业主代表为人数', message='请选择业主代表人数')
            return False
        if self.backexpertnum_var.get()=='' :
            messagebox.showwarning(title='注意递补专家为人数', message='请选择递补专家人数')
            return False
        return True

    def checktext(self):
        pass

    def showconfnum_func(self,*args):
        if self.checknum_func():
            self.expertnum_var.set(int(self.confnum_var.get())-int(self.bossnum_var.get())-self.selectedexpert)
            self.expertchoicenum_var.set(int(self.confnum_var.get())+int(self.backexpertnum_var.get())-int(self.bossnum_var.get())-self.selectedexpert)

    # 查看专家
    def showexpert_func(self):
        expertwindow(self)

class expertwindow():
    def __init__(self,mother):
        self.mother=mother
        self.top=Toplevel(self.mother.root,width=900, height=500)
        self.top.title('专家库')
        # self.top.attributes('-topmost', 1)

        # 基本信息
        self.expertnum_var=StringVar()
        Label(self.top, text='专家数量').grid(row=0, column=0)
        self.expertnum_et = Entry(self.top, state='readonly',textvariable=self.expertnum_var).grid(row=0, column=1)
        Button(self.top, text='导出数据', command=self.exportexpert_func, width=30).grid(row=0, column=2)
        Button(self.top,text='退出',command=self.top.destroy,width=30).grid(row=0,column=3)



        # 初始化树
        self.experttree = ttk.Treeview(self.top, show="headings", height=20, columns=("a", "b", "c", "d", "e", "f"))
        self.expertvbar = ttk.Scrollbar(self.top, orient=VERTICAL, command=self.experttree.yview)
        # 定义树形结构与滚动条
        self.experttree.configure(yscrollcommand=self.expertvbar.set)
        # 表格的标题
        self.experttree.column("a", width=80, anchor="center")
        self.experttree.column("b", width=50, anchor="center")
        self.experttree.column("c", width=200, anchor="center")
        self.experttree.column("d", width=200, anchor="center")
        self.experttree.column("e", width=180, anchor="center")
        self.experttree.column("f", width=200, anchor="center")
        self.experttree.heading("a", text="姓名")
        self.experttree.heading("b", text="性别")
        self.experttree.heading("c", text="现任职务")
        self.experttree.heading("d", text="工作单位")
        self.experttree.heading("e", text="电话")
        self.experttree.heading("f", text="身份证号")
        # 调用方法获取表格内容插入
        self.experttree.grid(row=1, column=0, sticky='NEW', columnspan=4)
        self.expertvbar.grid(row=1, column=4, sticky='NS', columnspan=4)


        # 显示专家树
        self.showexperttree_func()


    def __del__(self):
        pass

    # 显示专家树
    def showexperttree_func(self):

        try:
            self.dfexpert=pd.read_sql('select * from tbexpert',self.mother.dbconn)
            self.dfcata = pd.read_sql('select * from tbcata', self.mother.dbconn)
            self.expertnum_var.set(self.dfexpert['姓名'].count())
            for index, row in self.dfexpert.iterrows():
                self.experttree.insert("", "end", values=(row[1], row[2], row[3], row[4], row[8], row[80]))
        except:
            self.experttree.insert("", "end", values=('没有数据', '', '', '', '', ''))
    def exportexpert_func(self):

        try:
            writer = pd.ExcelWriter('导出专家库.xls')
            self.dfexpert.to_excel(writer,sheet_name='专家名单',index=False)
            self.dfcata.to_excel(writer, sheet_name='专业分类', index=False)
            writer.save()
            messagebox.showinfo(title='导出成功',message='导出成功，请在程序同一个文件夹查看。')

        except:
            messagebox.showwarning(title='导出失败',message='导出失败，请重新导入！')






if __name__ == '__main__':
     expertChoice()


