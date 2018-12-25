from tkinter import *
import tkinter.ttk

root=Tk()
root.title("测试")
root.resizable(0, 0)
root.geometry('1024x768')
for i in range(10):
    for j in range(0,10,2):
        if i==1 and j==4:
            Label(root,text='我只是想拉长一点\n换个行'+str(i)+str(j)).grid(row=i,column=j,rowspan=2,padx=10,pady=10)
            Entry(root).grid(row=i,column=j+1)
        else:
            Label(root, text=str(i) + str(j)).grid(row=i, column=j,padx=10)


mainloop()