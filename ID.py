from tkinter import*
import tkinter.messagebox as msg
import pyodbc
import manager
# import customer
import ManagerSystem
#import r_operation

# 檢查賬號
def id_check(a):    
    global id
    if a == '1':    # 在管理員界面下登錄，引數是1
    # 把賬號/密碼框框里輸入的字串賦值給id/password
        id = manager.entry_name.get()
        password = manager.entry_key.get()
    # else:           # 在讀者界面下登錄，引數是0
    #     id = customer.entry_name.get()
    #     password = customer.entry_key.get()
    getid()     # 最后得到id

    # connect access
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/tina0_000\Desktop/newDB/DB.accdb;')
    cursor = conn.cursor()
    sql = "SELECT passwd FROM user WHERE id='%s' AND job='%s'" % (id,a)
    cursor.execute(sql)     # 執行sql陳述句
    result = cursor.fetchone()      # 得到的結果回傳給result陣列

    if result:      # 有帳號
            if password == result[0]:
                success_login(a)    # 密碼對上了，進入對應的讀者/管理員操作界面
            else:                   # 有賬號但密碼沒對上
               msg._show(title='錯誤',message='賬號或密碼輸入錯誤！')
    else:           # 沒有賬號
        msg._show(title='錯誤！',message='您輸入的用戶不存在！請先注冊！')
        if a=='1':
            manager.root1.destroy()     # 關閉登錄小視窗，回到管理員界面
        # elif a=='0':
        #     customer.root1.destroy()
    conn.close()                        # 查詢完一定要關閉資料庫啊

def success_login(a):       # 成功登錄
    if a == '1':
        manager.root1.destroy()
        ManagerSystem.frame()     # 銷毀登錄注冊界面，跳轉到管理員的操作界面

    # elif a == '0':
    #     customer.root1.destroy()
    #     r_operation.frame()     # 銷毀登錄注冊界面，跳轉到讀者的操作界面

def id_write(a):                # 寫入（注冊）賬號
    conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/tina0_000\Desktop/newDB/DB.accdb;')
    cursor = conn.cursor()
    if a=='1':      # 跟check函式里邊十分類似
        id = manager.entry_name.get()           # 得到輸入的賬號
        password = manager.entry_key.get()      # 得到輸入的密碼
        confirm = manager.entry_confirm.get()   # 得到輸入的確認密碼
    # elif a=='0':
    #     id = customer.entry_name.get()
    #     password = customer.entry_key.get()
    #     confirm = customer.entry_confirm.get()

    sql0 = "SELECT id FROM user WHERE id='%s' AND job='%s'" % (id,a)
    sql1 = "INSERT INTO user(id, passwd, job) VALUES(?,?,?) "
    sql1_ = (id, password, a)

#首先檢查兩次輸入的密碼是否一致，一致后再檢查注冊的賬號是否已經存在
    if password == confirm:
        cursor.execute(sql0)
        result = cursor.fetchone()
        if result:
            msg.showerror(title='錯誤！', message='該賬號已被注冊，請重新輸入！')
        else:
            cursor.execute(sql1, sql1_)
            conn.commit()
            conn.close()
            msg.showinfo(title='成功！', message='注冊成功，請登錄！')

    else:
        msg.showerror(title='錯誤！', message='兩次密碼不一致，請重新輸入！')

def getid():
    return id