# 新增、刪除、修改、查詢、列印
import datetime
from datetime import date
from tkcalendar import DateEntry
from tkinter import*
from PIL import Image,ImageTk
from tkinter import ttk
import tkinter.messagebox as msg
import ManagerSystem
import pyodbc
import customerInfo

#初始登入畫面
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
    lable0 = Label(root, text='靜止戶管理', bg='pink', font=('微軟雅黑', 50)).pack()  # 上

    Button(root, text='新增', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=edit_CusInfo).place(x=400, y=250)
    Button(root, text='查詢', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=search_fixCusInfo).place(x=400, y=350)
    Button(root, text='退出', font=('微軟雅黑', 15), width=10, height=2, bg='yellow', command=enter_last).place(x=400, y=450)
    root.mainloop()

def edit_CusInfo():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('900x300')
    win.wm_attributes('-topmost', 1)
    lable1 = Label(win, text='新增靜止戶資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    global a_name
    Label(win, text='身分證字號/統一編號：', font=('宋體', 12)).place(x=30, y=150)
    a_name = Entry(win, font=('宋體', 12), width=12)#白色格子
    a_name.place(x=200, y=150)

    global b_name
    Label(win, text='客戶姓名：', font=('宋體', 12)).place(x=300, y=150)
    b_name = Entry(win, font=('宋體', 12), width=10)
    b_name.place(x=380, y=150)

    global birth
    Label(win, text='生日：', font=('宋體', 12)).place(x=450, y=150)
    tonow = datetime.datetime.now() #抓今天時間
    birth = DateEntry(win, width=10, year=tonow.year, month=tonow.month, day=tonow.day)
    birth.pack(padx=10, pady=10)
    birth.place(x=500, y=150)

    global phone
    Label(win, text='電話：', font=('宋體', 12)).place(x=30, y=195)
    phone = Entry(win, font=('宋體', 12), width=17)
    phone.place(x=80, y=195)

    global Email
    Label(win, text='Email：', font=('宋體', 12)).place(x=200, y=195)
    Email = Entry(win, font=('宋體', 12), width=25)
    Email.place(x=250, y=195)

    global coupon
    Label(win, text='會員折扣：', font=('宋體', 12)).place(x=450, y=195)
    coupon = Entry(win, font=('宋體', 12), width=5)
    coupon.place(x=530, y=195)

    Button(win, text='確認添加', font=('宋體', 12), width=10, command=add).place(x=650, y=220)
    
    global age
    age = 2
    getbirth = birth.get_date()
    print(getbirth)
    age=calculate_age(getbirth)
    print(age)

def add():#添加客戶資訊到資料庫中
    try:
        getbirth = birth.get_date()
        age=str(calculate_age(getbirth))
        sql="INSERT INTO 靜止戶 ([身分證字號/統一編號], [客戶姓名], [生日], [電話], [Email], [會員折扣], [年齡]) VALUES(?,?,?,?,?,?,?)"
        sql_= (a_name.get(),b_name.get(),birth.get(),phone.get(),Email.get(),coupon.get(),age)
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/tina0_000\Desktop/newDB/DB.accdb;')
        conn.cursor()#呼叫使用的意思
        cursor = conn.cursor()
        try:
            cursor.execute(sql, sql_)
            print('execute!')
        except Exception as ex:#更詳細的找錯誤原因
            print(ex)
        conn.commit()   # 這句不可或缺，當我們修改資料完成后必須要確認才能真正作用到資料庫里
        conn.close()
        msg.showinfo(title='成功！', message='已入庫！')
    except:
        msg.showinfo(title='失敗！', message='請填選完整！')


def add_fixedCusInfo():
    print(customerInfo.search_info)
    try:
        sql="INSERT INTO 靜止戶 ([身分證字號/統一編號], [客戶姓名], [生日], [電話], [Email], [會員折扣], [年齡]) VALUES(?,?,?,?,?,?,?)"    
        try:
            sql_= (customerInfo.search_info[0], customerInfo.search_info[1], customerInfo.search_info[2], customerInfo.search_info[3], customerInfo.search_info[4], customerInfo.search_info[5], customerInfo.search_info[6])    
        except Exception as ex:#更詳細的找錯誤原因
            print('sql_')
            print(ex)
        conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/tina0_000\Desktop/newDB/DB.accdb;')
        conn.cursor()#呼叫使用的意思   
        cursor = conn.cursor()
        cursor.execute(sql, sql_)
        conn.commit()   # 這句不可或缺，當我們修改資料完成后必須要確認才能真正作用到資料庫里
        conn.close()
        msg.showinfo(title='stop 成功！', message='已入庫！')
    except Exception as ex:
        msg.showinfo(title='失敗 stop！', message='請填選完整！')
        print(ex)
    print('d')         

def search_fixCusInfo():
    global win
    win = Tk()
    win.title('管理員')
    win.geometry('1000x500')
    lable1 = Label(win, text='查詢靜止戶資訊:', font=('微軟雅黑', 20)).place(x=30, y=50)

    global a_name
    Label(win, text='身分證字號/統一編號：', font=('宋體', 12)).place(x=30, y=150)
    a_name = Entry(win, font=('宋體', 12), width=12)#白色格子
    a_name.place(x=200, y=150)

    global b_name
    Label(win,text='客戶名稱：',font=('宋體',12)).place(x=30,y=195)
    b_name=Entry(win,font=('宋體', 12),width=10)
    b_name.place(x=110,y=195)

    Button(win,text='搜索',font=('宋體', 12), bg='yellow', width=10,command=search).place(x=600,y=195)
    
    global tree # 建立樹形圖
    yscrollbar = ttk.Scrollbar(win, orient='vertical')  # 滑軌
    tree = ttk.Treeview(win, columns=('1', '2', '3', '4', '5', '6', '7'), show="headings",yscrollcommand=yscrollbar.set)
    # 設定寬度、對其方式
    tree.column('1', width=100, anchor='center')
    tree.column('2', width=100, anchor='center')
    tree.column('3', width=100, anchor='center')
    tree.column('4', width=100, anchor='center')
    tree.column('5', width=100, anchor='center')
    tree.column('6', width=100, anchor='center')
    tree.column('7', width=100, anchor='center')

    # 設定欄位名稱
    tree.heading('1', text='身分證字號/統一編號')
    tree.heading('2', text='客戶姓名')
    tree.heading('3', text='生日')
    tree.heading('4', text='電話')
    tree.heading('5', text='Email')
    tree.heading('6', text='會員折扣')
    tree.heading('7', text='年齡')
    tree.place(x=150, y=250)

def search():
    clear_search()
    if a_name.get() == '':
        if b_name.get() == '全部':
          sql = "SELECT * FROM 靜止戶"
        else:
            try:
                sql = "SELECT * FROM 靜止戶 WHERE 客戶姓名='%s'"%(b_name.get())
                #Button(win,text='刪除',font=('宋體', 12), bg='yellow', width=10,command=add_fixedCusInfo).place(x=725,y=195)
                #Button(win,text='編輯',font=('宋體', 12), bg='yellow', width=10,command=update).place(x=725,y=195)
            except Exception as ex:#更詳細的找錯誤原因
                print(ex)
    else:
        if a_name.get() == '全部':
          sql = "SELECT * FROM 靜止戶"
        else: 
            try:
                sql = "SELECT * FROM 靜止戶 WHERE [身分證字號/統一編號]='%s'"%(a_name.get())
            except Exception as ex:#更詳細的找錯誤原因
                print(ex)

    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/tina0_000\Desktop/newDB/DB.accdb;')
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
        global search_info
        search_info = []
        
        for i in range(0,len(E2)):#查詢到的結果依次插入到表格中
            tree.insert('',i,values=(E2[i]))
        search_info = S
        print(search_info)

    else :
        tree.insert('', 0,values=('查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果','查詢不到結果'))

    conn.close()  

def update():#表示按下的按鈕是編輯鍵
    print('c')

def enter_last():# 退出管理員界面，跳轉至初始界面
    root.destroy()
    ManagerSystem.frame()

def calculate_age(getbirth):#用年分算年齡的
    if getbirth != date.today() :
        today = date.today()
        age1 = today.year - getbirth.year
        return age1
    else:
        return age

def clear_search():
    tree.delete(*tree.get_children())
