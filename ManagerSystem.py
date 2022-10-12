from tkinter import*
from tkinter import ttk
import tkinter.messagebox as msg
import pyodbc
import manager
import product
import customerInfo
import stop
import supplier
import order

def frame():
    global win
    win=Tk()
    win.title('東海花圃管理系統')
    screenWidth = win.winfo_screenwidth()      # 抓取螢幕寬度
    screenHeight = win.winfo_screenheight()    # 抓取螢幕高度
    w = 900                                     # 視窗寬
    h = 650                                     # 視窗高
    x = (screenWidth - w) / 2                   # 視窗左上角x軸位置
    y = (screenHeight - h ) / 2                 # 視窗左上角Y軸位置
    win.geometry("%dx%d+%d+%d" % (w,h,x,y))
    win.resizable(False, False)                # 限制視窗最大化
    lable0 = Label(win, text='管理員系統', bg='pink', font=('微軟雅黑', 50)).pack()  # 上

    Button(win, text='花草苗木', font=('微軟雅黑', 15), width=10, height=2,command=enter_product, bg='yellow').place(x=400, y=150)
    Button(win, text='客    戶', font=('微軟雅黑', 15), width=10, height=2,command=enter_customerInfo, bg='yellow').place(x=400, y=250)
    Button(win, text='靜 止 戶', font=('微軟雅黑', 15), width=10, height=2,command=enter_stop, bg='yellow').place(x=400, y=350)
    Button(win, text='供 應 商', font=('微軟雅黑', 15), width=10, height=2,command=enter_supplier, bg='yellow').place(x=400, y=450)
    Button(win, text='訂    單', font=('微軟雅黑', 15), width=10, height=2,command=enter_order, bg='yellow').place(x=400, y=550)
    win.mainloop()

def enter_product():
    win.destroy()
    product.frame()

def enter_customerInfo():
    win.destroy()
    customerInfo.frame()

def enter_stop():
    win.destroy()
    stop.frame()

def enter_supplier():
    win.destroy()
    supplier.frame()

def enter_order():
    win.destroy()
    order.frame()