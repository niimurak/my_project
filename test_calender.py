import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import os
from tkcalendar import Calendar, DateEntry
import pandas as pd
import datetime


# 出力処理
def click_exe_button():
    start = calender_date.get_date()
    #end = calender_date2.get_date().split("/")
    end = calender_date2.get_date()
    log.insert(END, '開始日'+ str(start) + '\n')
    #log.insert(END, '終了日'+'20'+end[2]+'-'+end[0]+'-'+end[1]+'\n')
    log.insert(END, '終了日'+ str(end) +'\n')
    dt = datetime.datetime.now()  # UTC
    str_dt = dt.strftime('%Y/%m/%d %H:%M:%S')
    log.insert(END, '処理完了' + str_dt + '\n')


if __name__ == '__main__':
    # ウィンドウを作成
    root = tkinter.Tk()
    root.title("アプリテスト") # アプリの名前
    root.geometry("480x500") # アプリの画面サイズ

    # Frame2の作成
    frame2 = ttk.Frame(root, padding=10)
    frame2.grid()
    
    # 日付選択ボタン
    
    start = StringVar()
    start.set('開始日')
    start_label = ttk.Label(frame2, textvariable=start)
    start_label.grid(row=0, column=0)
    calender_date =DateEntry(frame2)
    calender_date.grid(row=1, column=0)
    
    end = StringVar()
    end.set('終了日')
    end_label = ttk.Label(frame2, textvariable=end)
    end_label.grid(row=0, column=2)
    #calender_date2 = Calendar(frame2, date_patternstr="y-mm-dd")
    #calender_date2.grid(row=1, column=2)
    calender_date2 =DateEntry(frame2)
    calender_date2.grid(row=1, column=2)
    
    # Frame3の作成
    frame3 = ttk.Frame(root, padding=10)
    frame3.grid()
    
    # 処理ボタンの作成
    export_button = ttk.Button(frame3, text='日付確認', command=click_exe_button, width=20)
    export_button.grid(row=0, column=1)
    
    # Frame4の作成
    frame4 = ttk.Frame(root, padding=10)
    frame4.grid()
    
    # ログ表示BOX
    log = Text(frame4,width=50, height=12,borderwidth=5,wrap='none') 
    log.grid(row=1, column=1) 
    
    # ウィンドウを動かす
    root.mainloop()