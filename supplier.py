# 新增、查詢、列印

from tkinter import*
from tkinter import ttk
import tkinter.messagebox as msg
import ManagerSystem
import pyodbc
from tkcalendar import DateEntry
import datetime


def frame():
    global root
    root=Tk()
    root.title('東海花圃管理系統')
    screenWidth = root.winfo_screenwidth()      # 抓取螢幕寬度
    screenHeight = root.winfo_screenheight()    # 抓取螢幕高度
    w = 900                                     # 視窗寬
    h = 650                                     # 視窗高
    x = (screenWidth - w) / 2                   # 視窗左上角x軸位置
    y = (screenHeight - h ) / 2                 # 視窗左上角Y軸位置
    root.geometry("%dx%d+%d+%d" % (w,h,x,y))
    root.resizable(False, False)                # 限制視窗最大化
    lable0 = Label(root, text='供應商管理', bg='pink', font=('微軟雅黑', 50)).pack()  # 上

    Button(root, text='新增', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=insert_supplier).place(x=400, y=250)
    Button(root, text='查詢', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=search_supplier).place(x=400, y=350)
    Button(root, text='退出', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=enter_last).place(x=400, y=450)
    root.mainloop()

def insert_supplier():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('900x300')
    win.resizable(False, False)                # 限制視窗最大化
    lable1 = Label(win, text='新增供應商資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    global supplier_name
    Label(win, text='供應商名稱：', font=('宋體', 12)).place(x=30, y=150)
    supplier_name = Entry(win, font=('宋體', 12), width=15)
    supplier_name.place(x=130, y=150) 

    global supplier_num
    Label(win, text='統一編號：', font=('宋體', 12)).place(x=265, y=150)
    supplier_num = Entry(win, font=('宋體', 12), width=15)
    supplier_num.place(x=350, y=150) 

    global supplier_tel
    Label(win, text='電話：', font=('宋體', 12)).place(x=485, y=150)
    supplier_tel = Entry(win, font=('宋體', 12), width=15)
    supplier_tel.place(x=540, y=150) 

    global supplier_email
    Label(win, text='E-mail：', font=('宋體', 12)).place(x=30, y=195)
    supplier_email = Entry(win, font=('宋體', 12), width=20)
    supplier_email.place(x=90, y=195) 

    global supplier_person
    Label(win, text='負責人：', font=('宋體', 12)).place(x=265, y=195)
    supplier_person = Entry(win, font=('宋體', 12), width=10)
    supplier_person.place(x=335, y=195)

    Button(win, text='確認添加', font=('宋體', 12), width=10, command=add).place(x=700, y=200)


def add():  # 添加供應商資訊到資料庫中
    try:
        sql="INSERT INTO 供應商 (供應商統一編號, 供應商名稱, 電話, Email, 負責人) VALUES(?,?,?,?,?)"
        sql_= (supplier_num.get(), supplier_name.get(), supplier_tel.get(), supplier_email.get(), supplier_person.get())
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
        cursor = conn.cursor()
        cursor.execute(sql, sql_)
        conn.commit()   # 這句不可或缺，當我們修改資料完成后必須要確認才能真正作用到資料庫里
        conn.close()
        msg.showinfo(title='成功！', message='已入庫！')
    except:
        msg.showinfo(title='失敗！', message='請填選完整！')

def search_supplier():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('1200x700')
    win.resizable(False, False)                # 限制視窗最大化
    lable1 = Label(win, text='查詢供應商資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    global supplier_name
    Label(win,text='供應商名稱：',font=('宋體',12)).place(x=200,y=150)
    supplier_name=Entry(win,font=('宋體', 12),width=15)
    supplier_name.place(x=300,y=150)

    global supplier_email
    Label(win,text='E-mail：',font=('宋體',12)).place(x=450,y=150)
    supplier_email=Entry(win,font=('宋體', 12),width=15)
    supplier_email.place(x=510,y=150)

    Button(win,text='搜索',font=('宋體', 12), bg='yellow', width=10,command=search).place(x=900,y=150)

    global tree # 建立樹形圖
    yscrollbar = ttk.Scrollbar(win, orient='vertical')  # 滑軌
    tree = ttk.Treeview(win, columns=('1', '2', '3', '4', '5'), show="headings",yscrollcommand=yscrollbar.set)
    # 設定寬度、對其方式
    tree.column('1', width=150, anchor='center')
    tree.column('2', width=150, anchor='center')
    tree.column('3', width=150, anchor='center')
    tree.column('4', width=200, anchor='center')
    tree.column('5', width=150, anchor='center')
    # 設定欄位名稱
    tree.heading('1', text='編號')
    tree.heading('2', text='名稱')
    tree.heading('3', text='電話')
    tree.heading('4', text='Email')
    tree.heading('5', text='負責人')
    tree.place(x=200, y=250)
    yscrollbar.place(x=1000,y=250)

    
    Label(win,width=10,bg='yellow',text='',font=('宋體',12,'bold')).place(x=870,y=500)
    Label(win,text='筆',font=('宋體',12,'bold')).place(x=980,y=500)

def search():    #動態查詢
    clear_search()
    if supplier_name.get() == '全部':
        sql = "SELECT * FROM 供應商"
    else:
        sql = "SELECT * FROM 供應商 WHERE 供應商名稱='%s'"%(supplier_name.get())
    
    if supplier_email.get() != '':
        sql = "SELECT * FROM 供應商 WHERE Email='%s'"%(supplier_email.get())
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
    cursor = conn.cursor()
    cursor.execute(sql)
    results=cursor.fetchall()
    if results:
        E = []     # 存放未整理資料
        for row in results:
            E.append(row)
        E1 = []     # 小陣列
        E2 = []     # 大陣列
        for o in range(len(E)):
            S=E[o]
            for i in range(len(S)):
                try:
                    E1.append(str(S[i].date()))
                except:
                    E1.append(S[i])
            E2.append(E1)
            E1 = []

        for i in range(0,len(E2)):#查詢到的結果依次插入到表格中
            tree.insert('',i,values=(E2[i]))
    else :
        tree.insert('', 0,values=('查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果'))
    
    getTotall()

    conn.close()

def clear_search():
    tree.delete(*tree.get_children())


def getTotall():
    global totall
    totall = 0
    if supplier_name.get() == '全部':
        sql = "SELECT * FROM 供應商"
    else:
        sql = "SELECT * FROM 供應商 WHERE 供應商名稱='%s'"%(supplier_name.get())
    
    if supplier_email.get() != '':
        sql = "SELECT * FROM 供應商 WHERE Email='%s'"%(supplier_email.get())
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
    cursor = conn.cursor()
    cursor.execute(sql)
    results=cursor.fetchall()
    conn.close()
    for row in results:
        totall += 1
    Label(win,width=10,bg='yellow',text=totall,font=('宋體',12,'bold')).place(x=870,y=500)

def enter_last():# 退出管理員界面，跳轉至初始界面
    root.destroy()
    ManagerSystem.frame()
