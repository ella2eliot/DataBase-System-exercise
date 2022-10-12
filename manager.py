from tkinter import*
from PIL import Image,ImageTk
import tkinter.messagebox as msg #這個是會彈出一個警告/提示小框
import initial
import ID
import ManagerSystem

def frame():#管理員界面
    global root
    root= Tk()
    screenWidth = root.winfo_screenwidth()      # 抓取螢幕寬度
    screenHeight = root.winfo_screenheight()    # 抓取螢幕高度
    w = 900                                     # 視窗寬
    h = 650                                     # 視窗高
    x = (screenWidth - w) / 2                   # 視窗左上角x軸位置
    y = (screenHeight - h ) / 2                 # 視窗左上角Y軸位置
    root.resizable(False, False)                # 限制視窗最大化
    root.geometry("%dx%d+%d+%d" % (w,h,x,y))
    root.title('東海花圃')
    
    lable0 = Label(root, text='管理員登錄', bg='pink', font=('微軟雅黑', 50)).pack()  # 上

    # logo
    photo = Image.open("Thu_logo.png")
    photo = photo.resize((200,200))
    img = ImageTk.PhotoImage(photo)
    label=Label(root,text="Thu_logo",image=img)
    label.place(relx=0.3,rely=0.15,relwidth=0.4,relheight=0.4)

    #buttom
    Button(root, text='登錄', font=('微軟雅黑', 15), width=10, height=2, command=login, bg='yellow').place(x=200, y=500)
    Button(root, text='注冊', font=('微軟雅黑', 15), width=10, height=2, command=register, bg='yellow').place(x=400, y=500)
    Button(root, text='退出', font=('微軟雅黑', 15), width=10, height=2, command=exit_manager, bg='yellow').place(x=600, y=500)
    root.mainloop()

# 登錄小視窗
def login():
    global root1
    root1=Tk()
    root1.wm_attributes('-topmost', 1)  #將登錄視窗置頂不至于被遮到下面
    root1.title('管理員登錄')
    root1.geometry('500x300')
    root1.resizable(False, False)                # 限制視窗最大化

    lable1 = Label(root1, text='賬號：', font=25).place(x=100,y=50)
    lable2 = Label(root1, text='密碼：', font=25).place(x=100, y=100)

    global entry_name, entry_key
    name=StringVar()
    key = StringVar()

    entry_name = Entry(root1, textvariable=name, font=25)
    entry_name.place(x=180, y=50)
    entry_key = Entry(root1, textvariable=key, font=25,show='*')
    entry_key.place(x=180,y=100)
    # 百度：tkinter要求由按鈕（或者其它的插件）觸發的控制器函式不能含有引數,若要給函式傳遞引數，需要在函式前添加lambda：
    button1 = Button(root1, text='確定', height=2, width=10, command=lambda: ID.id_check('1'))
    button1.place(x=210, y=180)
#當我們輸入賬號和密碼，點擊確定時候，會呼叫ID模塊里的id_check()函式，1是引數，表示其身份是管理員

def register():#注冊小視窗
    global root2
    root2 = Tk()
    root2.wm_attributes('-topmost', 1)
    root2.title('管理員注冊')
    root2.geometry('500x300')
    root2.resizable(False, False)                # 限制視窗最大化

    lable1 = Label(root2, text='賬號：', font=25).place(x=100, y=50)
    lable2 = Label(root2, text='密碼：', font=25).place(x=100, y=100)
    lable2 = Label(root2, text='確認密碼：', font=25).place(x=80, y=150)

    global entry_name, entry_key, entry_confirm
    name = StringVar()
    key = StringVar()
    confirm = StringVar()
    entry_name = Entry(root2, textvariable=name, font=25)
    entry_name.place(x=180, y=50)
    entry_key = Entry(root2, textvariable=key, font=25, show='*')
    entry_key.place(x=180, y=100)
    entry_confirm = Entry(root2, textvariable=confirm,font=25, show='*')
    entry_confirm.place(x=180, y=150)
    # 百度：tkinter要求由按鈕（或者其它的插件）觸發的控制器函式不能含有引數,若要給函式傳遞引數，需要在函式前添加lambda：
    button1 = Button(root2, text='確定', height=2, width=10, command=lambda: ID.id_write('1'))
    button1.place(x=210, y=200)
#當我們點擊確定的時候，會呼叫ID模塊里的id_write()函式，1是引數，表示其身份是管理員

def exit_manager():# 退出管理員界面，跳轉至初始界面
    root.destroy()
    initial.frame()