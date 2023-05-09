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
    lable0 = Label(root, text='訂單管理', bg='pink', font=('微軟雅黑', 50)).pack()  # 上

    Button(root, text='新增', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=insert_order).place(x=400, y=250)
    Button(root, text='查詢', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=search_order).place(x=400, y=350)
    Button(root, text='退出', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=enter_last).place(x=400, y=450)
    root.mainloop()

def insert_order():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('900x300')
    win.resizable(False, False)                # 限制視窗最大化
    #win.wm_attributes('-topmost', 1)
    lable1 = Label(win, text='新增訂單資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    sql="SELECT [花草苗木名稱] FROM 產品"
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
    Label(win, text='植物名稱：', font=('宋體', 12)).place(x=365, y=150)
    cursor = conn.cursor()
    cursor.execute(sql)
    SE = []
    for row in cursor.fetchall():
        SE.append(row[0])
    global product_name
    comvalue = StringVar()
    product_name = ttk.Combobox(win, textvariable=comvalue, height=10, width=10)
    product_name.place(x=450, y=150)
    product_name['values'] = (SE)
    product_name.current(0)      # 默認顯示'全部'
    print(product_name.get())

    sql="SELECT [身分證字號/統一編號] FROM 客戶"
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
    Label(win, text='客戶身分證字號/統一編號：', font=('宋體', 12)).place(x=30, y=150)
    cursor = conn.cursor()
    cursor.execute(sql)
    SE = []
    for row in cursor.fetchall():
        SE.append(row[0])

    global person_num
    comvalue = StringVar()
    person_num = ttk.Combobox(win, textvariable=comvalue, height=10, width=10)
    person_num.place(x=230, y=150)
    person_num['values'] = (SE)
    person_num.current(0)      # 默認顯示'全部'
    print(person_num.get())

    # global person_num
    # Label(win, text='客戶身分證字號/統一編號：', font=('宋體', 12)).place(x=30, y=150)
    # person_num = Entry(win, font=('宋體', 12), width=15)
    # person_num.place(x=230, y=150) 
    
    # global product_name
    # Label(win, text='植物名稱：', font=('宋體', 12)).place(x=365, y=150)
    # product_name = Entry(win, font=('宋體', 12), width=10)
    # product_name.place(x=450, y=150) 

    global amount
    Label(win, text='數量：', font=('宋體', 12)).place(x=545, y=150)
    amount = Entry(win, font=('宋體', 12), width=5)
    amount.place(x=600, y=150)

    global price
    Label(win, text='售價：', font=('宋體', 12)).place(x=655, y=150)
    price = Entry(win, font=('宋體', 12), width=5)
    price.place(x=710, y=150)

    global order_date
    Label(win, text='訂購日期：', font=('宋體', 12)).place(x=30, y=195)
    tonow = datetime.datetime.now()
    order_date = DateEntry(win, width=10, year=tonow.year, month=tonow.month, day=tonow.day)
    order_date.pack(padx=10, pady=10)
    order_date.place(x=115, y=195)

    global ED_date  
    Label(win, text='預計交貨日期：', font=('宋體', 12)).place(x=220, y=195)
    tonow = datetime.datetime.now()
    ED_date = DateEntry(win, width=10, year=tonow.year, month=tonow.month, day=tonow.day)
    ED_date.pack(padx=10, pady=10)
    ED_date.place(x=335, y=195)

    global AD_date  
    Label(win, text='實際交貨日期：', font=('宋體', 12)).place(x=440, y=195)
    tonow = datetime.datetime.now()
    AD_date = DateEntry(win, width=10, year=tonow.year, month=tonow.month, day=tonow.day)
    AD_date.pack(padx=10, pady=10)
    AD_date.place(x=555, y=195)

    Button(win, text='確認添加', font=('宋體', 12), width=10, command=add).place(x=700, y=200)

def add():  # 添加資訊到資料庫中
    try:
        sql = "SELECT 花草苗木編號 FROM 產品 WHERE 花草苗木名稱='%s'"%(product_name.get())
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
        cursor = conn.cursor()
        cursor.execute(sql)
        results=cursor.fetchall()
        product_num=results[0]
        conn.close()
  
        sql = "SELECT 供應商名稱 FROM 產品 WHERE 花草苗木名稱='%s'"%(product_name.get())
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
        cursor = conn.cursor()
        cursor.execute(sql)
        results=cursor.fetchall()
        supplier_name=results[0]
        conn.close()
    
        totall_price = int(amount.get()) * int(price.get())

        sql = "SELECT [會員折扣] FROM [客戶] WHERE [身分證字號/統一編號]='%s'"%(person_num.get())
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
        cursor = conn.cursor()
        cursor.execute(sql)
        results=cursor.fetchall()
        person_off=results[0]
        conn.close()

        afteroff = int(totall_price * person_off[0])
    except:
        next

    r1 = order_date.get() <= AD_date.get()
    r2 = order_date.get() <= ED_date.get()
    if r1==True and r2==True:
        try:
            sql="INSERT INTO 客戶購買 ([花草苗木編號], [客戶身分證字號/統一編號], [花草苗木名稱], [供應商名稱], [購買數量], [售價], [總金額], [折扣後金額], [訂購日期], [預計交貨日期], [實際交貨日期]) VALUES(?,?,?,?,?,?,?,?,?,?,?)"
            sql_= (product_num[0], person_num.get(), product_name.get(), supplier_name[0], amount.get(), price.get(), totall_price, afteroff, order_date.get(), ED_date.get(), AD_date.get())
            conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
            cursor = conn.cursor()
            cursor.execute(sql, sql_)   #error
            conn.commit()   # 這句不可或缺，當我們修改資料完成后必須要確認才能真正作用到資料庫里
            conn.close()
            msg.showinfo(title='成功！', message='已入庫！')
        except Exception as ex:#更詳細的找錯誤原因
            print(sql_)
            print(ex)
            msg.showinfo(title='失敗！', message='請填選完整！')
    else:
        msg.showinfo(title='失敗！', message='預計交貨日期和實際交貨日期不正確！')


######進入查詢的部分

def search_order():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('1300x700')
    win.resizable(False, False)                # 限制視窗最大化
    #win.wm_attributes('-topmost', 1)
    lable1 = Label(win, text='查詢訂單資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    # global person_num
    # Label(win,text='客戶身分證字號/統一編號：',font=('宋體',12)).place(x=215,y=150)
    # person_num=Entry(win,font=('宋體', 12),width=15)
    # person_num.place(x=415,y=150)

    sql="SELECT [身分證字號/統一編號] FROM 客戶"
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
    Label(win, text='客戶身分證字號/統一編號：', font=('宋體', 12)).place(x=215, y=150)
    cursor = conn.cursor()
    cursor.execute(sql)
    SE = []
    for row in cursor.fetchall():
        SE.append(row[0])

    global person_num
    comvalue = StringVar()
    person_num = ttk.Combobox(win, textvariable=comvalue, height=10, width=10)
    person_num.place(x=415, y=150)
    person_num['values'] = (SE)
    person_num.current(0)      # 默認顯示'全部'
    print(person_num.get())

    Button(win,text='搜索',font=('宋體', 12), bg='yellow', width=10,command=search).place(x=1000,y=150)

    global tree # 建立樹形圖
    yscrollbar = ttk.Scrollbar(win, orient='vertical')  # 滑軌
    tree = ttk.Treeview(win, columns=('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'), show="headings",yscrollcommand=yscrollbar.set)
    # 設定寬度、對其方式
    tree.column('1', width=100, anchor='center')
    tree.column('2', width=100, anchor='center')
    tree.column('3', width=100, anchor='center')
    tree.column('4', width=100, anchor='center')
    tree.column('5', width=100, anchor='center')
    tree.column('6', width=100, anchor='center')
    tree.column('7', width=100, anchor='center')
    tree.column('8', width=100, anchor='center')
    tree.column('9', width=100, anchor='center')
    tree.column('10', width=100, anchor='center')
    tree.column('11', width=100, anchor='center')
    tree.column('12', width=100, anchor='center')
    # 設定欄位名稱
    tree.heading('1', text='訂單編號')
    tree.heading('2', text='產品編號')
    tree.heading('3', text='客戶身分證字號/統一編號')
    tree.heading('4', text='產品名稱')
    tree.heading('5', text='供應商名稱')
    tree.heading('6', text='購買數量')
    tree.heading('7', text='售價')
    tree.heading('8', text='總金額')
    tree.heading('9', text='折扣號金額')
    tree.heading('10', text='訂購日期')
    tree.heading('11', text='預計交貨日期')
    tree.heading('12', text='實際交貨日期')
    tree.place(x=30, y=250)

    #global totall
    Label(win,text='總計：',font=('宋體',12,'bold')).place(x=945,y=500)
    Label(win,width=10,bg='yellow',text='',font=('宋體',12,'bold')).place(x=1000,y=500)

def search():    #動態查詢
    clear_search()
    if person_num.get() == '全部':
        sql = "SELECT * FROM [客戶購買]"
    else:
        sql = "SELECT * FROM [客戶購買] WHERE [客戶身分證字號/統一編號]='%s'"%(person_num.get())
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
        tree.insert('', 0,values=('查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果'))
    
    getTotall()

    conn.close()

def clear_search():
    tree.delete(*tree.get_children())

def enter_last():# 退出管理員界面，跳轉至初始界面
    root.destroy()
    ManagerSystem.frame()

def getTotall():
    global totall
    totall = 0
    if person_num.get() == '全部':
        sql = "SELECT [總金額] FROM [客戶購買]"
    else:
        sql = "SELECT [總金額]  FROM [客戶購買] WHERE [客戶身分證字號/統一編號]='%s'"%(person_num.get())
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=../DB.accdb;')
    cursor = conn.cursor()
    cursor.execute(sql)
    results=cursor.fetchall()
    conn.close()
    E = []
    for row in results:
        E.append(row[0])
        totall += int(row[0])
    Label(win,width=10,bg='yellow',text=totall,font=('宋體',12,'bold')).place(x=1000,y=500)
    
