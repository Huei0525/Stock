from StockUtils import StockUtils
from datetime import datetime, timedelta # datetime:處理日期及時間; timedelta: 計算二個日期或時間的差異
import twstock     # 台灣股市 Python 庫
import talib       # 技術分析 Python 庫
import numpy as np # 科學計算庫
import yfinance as yf
import os
import pygsheets              # Google Sheets操作
import pandas as pd           # 數據分析和處理庫
import requests               # HTTP 請求庫
from bs4 import BeautifulSoup # HTML 和 XML 解析庫
import tkinter as tk # 標準 GUI 函式庫，用於建立視窗、按鈕、文字方塊等圖形使用者介面元素
# ttk 是 tkinter 的子模組，提供了一些更現代和主題化的控件，如按鈕、標籤、框架等
# filedialog 是 tkinter 中的一個模組，用於開啟檔案對話框，讓使用者選擇檔案或目錄
# messagebox 是 tkinter 中的一個模組，用於顯示訊息框，如資訊、警告、錯誤提示等
from tkinter import ttk, filedialog, messagebox
from tkinter import font
from tkcalendar import Calendar #為 tkinter 提供日曆零件


# new class
objStockUtils = StockUtils()
stockCode = "" # 股票代號
stockName = "" # 股票名稱

def showStockBasic(event=None):
    """
    依User輸入「股票代號/名稱」取得對應的「股票代號」及「股票名稱」，並更新於畫面上
    """
    print('\t ===== showStockBasic() ===== ')
    global stockCode, stockName
    userInput = inputStock.get().strip()
    # print(F"\t userInput= {type(userInput)} {userInput}")


    for code, info in twstock.codes.items():
        #  code= <class 'str'> 6197
        #  info= <class 'twstock.codes.codes.StockCodeInfo'> StockCodeInfo(type='股票', code='6197', name='佳必琪', ISIN='TW0006197007', start='2004/11/08', market='上市', group='電子零組件業', CFI='ESVUFR')
        if userInput == code or userInput == info.name:
            stockCode = code
            stockName = info.name
    print(F"\t [showStockBasic]stockCode = {type(stockCode)} {stockCode}")
    print(F"\t [showStockBasic]stockName = {type(stockName)} {stockName}")

    if stockCode and stockName:
        labelStockBasic.config(text=f"股票代號：{stockCode}, 股票名稱：{stockName}")
    else:
        labelStockBasic.config(text="找不到對應的股票資料，請確認")


def pick_date(entry):
    """
    使用 Tkinter 和 tkcalendar 庫來實現選擇日期的功能
    """
    def on_date_selected(event):
        """
        處理當用戶從日曆選擇日期時的事件
        """
        date = cal.selection_get() #返回當前日曆 cal 中選擇的日期
        entry.set(date.strftime('%Y-%m-%d')) #將選擇的日期格式化為 %Y-%m-%d 的字串並設置為 entry 中顯示的值
        date_window.destroy() #關閉日期選擇窗口

    #創建了一個新的頂層窗口 date_window，用來顯示日曆和日期選擇器
    date_window = tk.Toplevel(root)
    #創建了一個 Calendar 小部件 cal，放置在 date_window 中
    cal = Calendar(date_window, selectmode='day', date_pattern='yyyy-mm-dd', locale='zh_TW')
    #將日曆小部件放置在 date_window 中，設置了 padx 和 pady 作為內邊距
    cal.pack(padx=10, pady=10)
    #綁定了日曆 cal 的 <<CalendarSelected>> 事件到 on_date_selected 函數
    #當用戶選擇日期時，將觸發 on_date_selected 函數，進行後續的日期設置和窗口關閉操作
    cal.bind("<<CalendarSelected>>", on_date_selected)


def getStockData(event=None):
    """
    依User輸入「股票代號/名稱」取得對應的技術指標相關資訊
    """
    print('\t ===== getStockData() ===== ')
    global stockCode, stockName
    userInput = inputStock.get().strip()
    print(F"\t userInput= {type(userInput)} {userInput}")

    # 如果還沒有取到「股票代號/名稱」，重新取
    if not stockCode or not stockName:
        showStockBasic()
    print(F"\t [getStockData]stockCode = {type(stockCode)} {stockCode}")
    print(F"\t [getStockData]stockName = {type(stockName)} {stockName}")

    # get 輸入的「開始日期」、「結束日期」
    sStartDate = inputStartDate.get()
    sEndDate   = inputEndDate.get()
    print(F"\t [getStockData]sStartDate = {type(sStartDate)} {sStartDate}")
    print(F"\t [getStockData]sEndDate = {type(sEndDate)} {sEndDate}")

    try:
        dStartDate = datetime.strptime(sStartDate, '%Y-%m-%d')
        dEndDate   = datetime.strptime(sEndDate, '%Y-%m-%d')
    except ValueError:
        messagebox.showinfo("Error", "日期格式錯誤，請輸入 YYYY-MM-DD 格式")
        return

    if dStartDate >= dEndDate:
        messagebox.showinfo("Error", "「開始日期」必須早於「結束日期」")
        return

    # 取得各技術指標
    result = []
    objStockUtils = StockUtils()
    listStockInfo = objStockUtils.getStockAnalyze(stockCode,sEndDate,sStartDate)
    dictDays = objStockUtils.countIncreaseDaysHistory(listStockInfo)
    print(F"[upBuyingStock] listStockInfo size = {type(listStockInfo)} {len(listStockInfo)}")

    # 將 List 用日期降序排序
    listStockInfo.sort(key=lambda x: x[0], reverse=False)
    for a in range(len(listStockInfo)):
        listStock = listStockInfo[a]
        stockDate   = listStock[0] # 資料日期
        stockCode   = listStock[1] # 股票代號
        stockName   = listStock[2] # 股票名稱
        fClosePrice = listStock[3] # 收盤價
        iUpQty      = listStock[4] # 多頭數量
        iDownQty    = listStock[5] # 空頭數量
        iUpDays = dictDays[stockDate+"Up"]
        iDownDays = dictDays[stockDate+"Down"]
        print(F"\t [countIncreaseDays] a = {a} ==================================================")
        print(f"\t iUpDays = {iUpDays}")
        print(f"\t iDownDays = {iDownDays}")


        result.append([stockDate, stockCode, stockName, fClosePrice, iUpQty, iDownQty, iUpDays, iDownDays])

    header = f"股票代號: {stockCode}\n股票名稱: {stockName}\n日期範圍: {sStartDate} 至 {sEndDate}\n\n"
    header = f"{header} 第一筆的「連續多頭增加」及「連續空頭增加」因為沒有上一筆，故一定為0\n\n"

    # 將 textResult 文的狀態設置為 NORMAL ，允許修改內容
    textResult.config(state=tk.NORMAL)
    # 刪除 textResult 全部內容
    textResult.delete(1.0, tk.END)

    # 設置為微軟正黑體
    microsoft_zh_font = font.Font(family="Microsoft JhengHei", size=10)
    textResult.configure(font=microsoft_zh_font)

    # 在 textResult 文本小部件的末尾位置插入新文本。插入的文本由 header（表头信息）和结果数据 result 组成
    # textResult.insert(tk.END, header + '\n')
    textResult.insert(tk.END, header + '\n'.join([f"日期: {r[0]}, 收盤價: {r[3]:.2f}, 多頭指標: {r[4]}, 空頭指標: {r[5]}, 連續多頭增加: {r[6]}, 連續空頭增加: {r[7]}" for r in result]))

    # 設置 sStartDate 和 sEndDate 顏色
    start_index = textResult.search(sStartDate, "1.0", tk.END)
    end_index = textResult.search(sEndDate, "1.0", tk.END)
    if start_index:
        start_end_index = f"{start_index} + {len(sStartDate)}c"
        textResult.tag_add("blue", start_index, start_end_index)
    if end_index:
        end_end_index = f"{end_index} + {len(sEndDate)}c"
        textResult.tag_add("blue", end_index, end_end_index)
    textResult.tag_config("blue", foreground="blue")

    # 將 textResult 文的狀態設置為 DISABLED ，禁止修改內容
    textResult.config(state=tk.DISABLED)

    if exportExcel.get():
        sPath = inputFolderPath.get()
        print(f'\t 匯出Excel = {sPath}')
        objStockUtils.saveToExcel(sPath, listStockInfo, dictDays)
        messagebox.showinfo("Msg", "已查詢完成！並保存Excel")
    else:
        messagebox.showinfo("Msg", "已查詢完成！")

# print("==================================================")
root = tk.Tk() # 建立 tkinter 視窗物件
root.title("股票技術指標查詢") # 設定標題
root.resizable(False, False)  # 使窗口大小固定


# ------------------------------
# GUI-股票代號/名稱 (row=0)
# ------------------------------
ttk.Label(root, text="股票代號/名稱").grid(column=0, row=0, padx=10, pady=5, sticky=tk.W)
inputStock = ttk.Entry(root, width=20) #創建了一個輸入框
inputStock.insert(0, "佳必琪")
inputStock.grid(column=1, row=0, padx=10, pady=5, sticky=tk.W)

# 綁定 Enter 鍵事件
inputStock.bind("<Return>", showStockBasic)
# 綁定離開文字框事件
inputStock.bind("<FocusOut>", showStockBasic)
labelStockBasic = ttk.Label(root, text="")
labelStockBasic.grid(column=2, row=0, padx=50, pady=5, sticky=tk.W)


# ------------------------------
# GUI-開始日期 (row=1)
# ------------------------------
ttk.Label(root, text="開始日期").grid(column=0, row=1, padx=10, pady=5, sticky=tk.W)

# Set 預設日期: 
listDate = objStockUtils.getStartDate(datetime.today(),"")
# startDate = tk.StringVar(value=listDate[0].strftime('%Y-%m-%d'))
startDate = tk.StringVar(value="2024-06-26")
inputStartDate = ttk.Entry(root, textvariable=startDate, width=10)
inputStartDate.grid(column=1, row=1, padx=10, pady=5, sticky=tk.W)

buttonStartDate = ttk.Button(root, text="選擇日期", command=lambda: pick_date(startDate))
buttonStartDate.grid(column=2, row=1, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-結束日期 (row=2)
# ------------------------------
ttk.Label(root, text="結束日期").grid(column=0, row=2, padx=10, pady=5, sticky=tk.W)
# Set 預設日期
endDate = tk.StringVar(value=datetime.today().strftime('%Y-%m-%d'))
inputEndDate = ttk.Entry(root, textvariable=endDate, width=10)
inputEndDate.grid(column=1, row=2, padx=10, pady=5, sticky=tk.W)

buttonEndDate = ttk.Button(root, text="選擇日期", command=lambda: pick_date(endDate))
buttonEndDate.grid(column=2, row=2, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-匯出Excel勾選框 (row=3)
# ------------------------------
exportExcel = tk.BooleanVar()
checkExportExcel = ttk.Checkbutton(root, text="匯出Excel", variable=exportExcel)
checkExportExcel.grid(column=0, row=3, padx=10, pady=5, sticky=tk.W)

# 創建資料夾路徑輸入框
ttk.Label(root, text="資料夾路徑").grid(column=1, row=3, padx=10, pady=5)
# Set 預設路徑
defaultPath = tk.StringVar(value="C:/Users/user/Desktop/股票技術指標輸出")
inputFolderPath = ttk.Entry(root, textvariable=defaultPath, width=40)
inputFolderPath.grid(column=2, row=3, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-查詢按鈕 (row=4)
# ------------------------------
buttonQuery = ttk.Button(root, text="查詢", command=getStockData)
buttonQuery.grid(column=0, row=4, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-結果顯示
# ------------------------------
result_text = tk.StringVar()
# 加入 Frame 框架
#     sticky=(tk.W, tk.E, tk.N, tk.S)) : 使得框架可以隨窗口大小改變而伸展
frameResult = ttk.Frame(root)
frameResult.grid(column=0, row=5, columnspan=3, padx=10, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S))

# 創建了一個文本框，並將其放置在 frameResult 框架中
#     wrap='word' : 設置文本框文字換行方式為單詞換行，即在單詞邊界換行
#     state=tk.DISABLED :　設置文本框為禁用狀態，即用戶無法編輯
textResult = tk.Text(frameResult, wrap='word', state=tk.DISABLED, width=80, height=32)
textResult.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# 創建了一個垂直方向的捲動條 scrollbar，並將其連接到 frameResult 的垂直滾動命令
#     sticky=(tk.N, tk.S : 填滿垂直空間
scrollbar = ttk.Scrollbar(frameResult, orient=tk.VERTICAL, command=textResult.yview)
scrollbar.grid(column=1, row=0, sticky=(tk.N, tk.S))
# 當用戶拖動捲動條時，文本框的滾動位置會隨之改變
textResult['yscrollcommand'] = scrollbar.set


# ------------------------------
# GUI-按 Enter ，觸發查詢Method
# ------------------------------
root.bind('<Return>', getStockData)


# ------------------------------
# 使用 mainloop() 將其放在主迴圈中一直執行，直到使用者關閉該視窗才會停止運作
# ------------------------------
root.mainloop()