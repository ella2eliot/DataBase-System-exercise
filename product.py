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
    lable0 = Label(root, text='花草苗木管理', bg='pink', font=('微軟雅黑', 50)).pack()  # 上

    Button(root, text='新增', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=insert_product).place(x=400, y=250)
    Button(root, text='查詢', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=search_product).place(x=400, y=350)
    Button(root, text='退出', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=enter_last).place(x=400, y=450)
    root.mainloop()

def insert_product():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('900x300')
    win.resizable(False, False)                # 限制視窗最大化
    #win.wm_attributes('-topmost', 1)
    lable1 = Label(win, text='新增花草苗木資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    global product_name
    Label(win, text='植物名稱：', font=('宋體', 12)).place(x=30, y=150)
    product_name = Entry(win, font=('宋體', 12), width=10)
    product_name.place(x=115, y=150) 
    
    sql="SELECT 供應商名稱 FROM 供應商"
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/盈琪作業區/大三/資料庫/資料庫project/GardenShop/DB.accdb;')
    Label(win, text='供應商名稱：', font=('宋體', 12)).place(x=210, y=150)
    cursor = conn.cursor()
    cursor.execute(sql)
    SE = []
    for row in cursor.fetchall():
        SE.append(row[0])

    global supplier_name
    comvalue = StringVar()
    supplier_name = ttk.Combobox(win, textvariable=comvalue, height=10, width=10)
    supplier_name.place(x=310, y=150)
    supplier_name['values'] = (SE)
    supplier_name.current(0)      # 默認顯示'全部'

    global amount
    Label(win, text='數量：', font=('宋體', 12)).place(x=410, y=150)
    amount = Entry(win, font=('宋體', 12), width=5)
    amount.place(x=465, y=150)

    Label(win, text='單位：', font=('宋體', 12)).place(x=520, y=150)
    global unit_list     # 這個是一個下拉頁表項，只能從下面的list['values']里邊選
    comvalue = StringVar()
    unit_list = ttk.Combobox(win, textvariable=comvalue, height=10, width=5)
    unit_list.place(x=575, y=150)
    unit_list['values'] = ('束', '盆')
    unit_list.current(0)      # 默認顯示'全部'

    global price
    Label(win, text='單價：', font=('宋體', 12)).place(x=645, y=150)
    price = Entry(win, font=('宋體', 12), width=5)
    price.place(x=695, y=150)

    Label(win, text='存放位置：', font=('宋體', 12)).place(x=30, y=195)
    global pos_list     # 這個是一個下拉頁表項，只能從下面的list['values']里邊選
    comvalue = StringVar()
    pos_list = ttk.Combobox(win, textvariable=comvalue, height=10, width=9)
    pos_list.place(x=115, y=195)
    pos_list['values'] = ('店面', '二樓花房', '三樓花房', '四樓花房', '倉庫')
    pos_list.current(0)      # 默認顯示'全部'

    global cal
    Label(win, text='進貨日期：', font=('宋體', 12)).place(x=225, y=195)
    tonow = datetime.datetime.now()
    cal = DateEntry(win, width=10, year=tonow.year, month=tonow.month, day=tonow.day)
    cal.pack(padx=10, pady=10)
    cal.place(x=310, y=195)

    Button(win, text='確認添加', font=('宋體', 12), width=10, command=add).place(x=700, y=200)

def add():  # 添加花草苗木資訊到資料庫中
    try:
        sql="INSERT INTO 產品 (花草苗木名稱, 供應商名稱, 公司內現有數量, 單位, 單價, 公司內存放位置, 進貨日期) VALUES(?,?,?,?,?,?,?)"
        sql_= (product_name.get(), supplier_name.get(), amount.get(), unit_list.get(), price.get(), pos_list.get(), cal.get())
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/盈琪作業區/大三/資料庫/資料庫project/GardenShop/DB.accdb;')
        cursor = conn.cursor()
        cursor.execute(sql, sql_)
        conn.commit()   # 這句不可或缺，當我們修改資料完成后必須要確認才能真正作用到資料庫里
        conn.close()
        msg.showinfo(title='成功！', message='已入庫！')
    except:
        msg.showinfo(title='失敗！', message='請填選完整！')


def search_product():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('1200x700')
    win.resizable(False, False)                # 限制視窗最大化
    lable1 = Label(win, text='查詢花草苗木資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    global product_name
    Label(win,text='植物名稱：',font=('宋體',12)).place(x=215,y=150)
    product_name=Entry(win,font=('宋體', 12),width=15)
    product_name.place(x=300,y=150)

    Button(win,text='搜索',font=('宋體', 12), bg='yellow', width=10,command=search).place(x=900,y=150)

    global tree # 建立樹形圖
    yscrollbar = ttk.Scrollbar(win, orient='vertical')  # 滑軌
    tree = ttk.Treeview(win, columns=('1', '2', '3', '4', '5', '6', '7', '8', '9'), show="headings",yscrollcommand=yscrollbar.set)
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
    # 設定欄位名稱
    tree.heading('1', text='編號')
    tree.heading('2', text='名稱')
    tree.heading('3', text='供應商')
    tree.heading('4', text='數量')
    tree.heading('5', text='單位')
    tree.heading('6', text='單價')
    tree.heading('7', text='位置')
    tree.heading('8', text='進貨日期')
    tree.heading('9', text='小計')
    tree.place(x=150, y=250)

    #global totall
    Label(win,text='總計：',font=('宋體',12,'bold')).place(x=890,y=500)
    Label(win,width=10,bg='yellow',text='',font=('宋體',12,'bold')).place(x=945,y=500)

def search():    #動態查詢
    clear_search()
    if product_name.get() == '全部':
        sql = "SELECT * FROM 產品"
    else:
        sql = "SELECT * FROM 產品 WHERE 花草苗木名稱='%s'"%(product_name.get())
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/盈琪作業區/大三/資料庫/資料庫project/GardenShop/DB.accdb;')
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

def enter_last():# 退出管理員界面，跳轉至初始界面
    root.destroy()
    ManagerSystem.frame()

def getTotall():
    global totall
    totall = 0
    if product_name.get() == '全部':
        sql = "SELECT 小計 FROM 產品"
    else:
        sql = "SELECT 小計 FROM 產品 WHERE 花草苗木名稱='%s'"%(product_name.get())
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/盈琪作業區/大三/資料庫/資料庫project/GardenShop/DB.accdb;')
    cursor = conn.cursor()
    cursor.execute(sql)
    results=cursor.fetchall()
    conn.close()
    E = []
    for row in results:
        E.append(row[0])
        totall += row[0]
    Label(win,width=10,bg='yellow',text=totall,font=('宋體',12,'bold')).place(x=945,y=500)
    