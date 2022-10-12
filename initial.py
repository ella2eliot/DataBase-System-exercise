from tkinter import*
from PIL import Image,ImageTk
import manager

def frame():#初始界面
    global root
    root=Tk()
    screenWidth = root.winfo_screenwidth()      # 抓取螢幕寬度
    screenHeight = root.winfo_screenheight()    # 抓取螢幕高度
    w = 900                                     # 視窗寬
    h = 650                                     # 視窗高
    x = (screenWidth - w) / 2                   # 視窗左上角x軸位置
    y = (screenHeight - h ) / 2                 # 視窗左上角Y軸位置
    root.geometry("%dx%d+%d+%d" % (w,h,x,y))
    root.resizable(False, False)                # 限制視窗最大化
    # 歡迎詞
    root.title('東海花圃')
    lable0=Label(root,text='歡迎來到東海花圃',bg='pink',font=('微軟雅黑',50)).pack()
	
    # logo
    photo = Image.open("Thu_logo.png")
    photo = photo.resize((200,200))
    img = ImageTk.PhotoImage(photo)
    label=Label(root,text="Thu_logo",image=img)
    label.place(relx=0.3,rely=0.15,relwidth=0.4,relheight=0.4)

    # user choice
    # lable1=Label(root,text='請選擇用戶型別：',font=('微軟雅黑',20)).place(x=150,y=480)
    
    Button(root, text='管理員',font=('微軟雅黑',15),width=10, height=2,command=exit_manager,bg='yellow').place(x=400, y=500)

    root.mainloop()

# def exit_customer():# 跳轉至客戶界面    root.destroy()
#     customer.frame()

def exit_manager():# 跳轉至管理員界面
    root.destroy()
    manager.frame()

if __name__ == '__main__':
    frame()