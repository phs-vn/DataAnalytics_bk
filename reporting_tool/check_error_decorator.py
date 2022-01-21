from reporting_tool.trading_service.thanhtoanbutru import *
from tkinter import *


def Error_Handler(func):
    def Inner_Function(*args,**kwargs):
        try:
            func(*args,**kwargs)
            popupRoot = Tk()
            popupRoot.geometry('300x200+1290+628')
            msg = Label(popupRoot, text='Report run successfully!!!', font=15)
            msg.place(x=150, y=75, anchor='center')
            exit_button = Button(popupRoot, text="Exit", command=popupRoot.destroy)
            exit_button.place(x=150, y=110, anchor='center', width=50, height=30)
            popupRoot.mainloop()

        except (Exception,):
            err = f"The error of {func.__name__} is error!!!"
            # Gửi lỗi qua Mail
            outlook = Dispatch('outlook.application')
            mail = outlook.CreateItem(0)

            mail.To = 'huynam177199@gmail.com'
            mail.Subject = "'Hệ thống chạy các Báo cáo Tự động' đang xảy ra lỗi!!"
            mail.Body = err

            mail.Send()
    return Inner_Function


@Error_Handler
def Mean(a, b):
    print((a + b) / 2)


@Error_Handler
def Square(sq):
    print(sq * sq)


@Error_Handler
def Divide(l, b):
    print(b / l)


Mean(4, 5)
Square(14)
Divide(8, 4)
Square("three")
Divide("two", "one")
Mean("six", "five")
