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


def showStockBasic(inputStockCode, inputStockName):
    """
    依User輸入「股票代號/名稱」取得對應的「股票代號」及「股票名稱」，並更新於畫面上
    """
    print('\t ===== showStockBasic() ===== ')
    global stockCode, stockName

    # get User輸入「股票代號/名稱」
    sInStockCode = inputStockCode.get() if inputStockCode else ""
    sInStockName = inputStockName.get() if inputStockName else ""
    # 去除兩端空白字符
    sInStockCode = sInStockCode.strip()
    sInStockName = sInStockName.strip()
    print(F"\t sInStockCode= {type(sInStockCode)} {sInStockCode}")
    print(F"\t sInStockName= {type(sInStockName)} {sInStockName}")

    # listStockCode = []
    # listStockName = []

    if sInStockCode != "":
        # 原本想改可以輸入多隻股票，但有點麻煩，先註解
        # codes = sInStockCode.split(',')
        # for code in codes:
        #     code = code.strip()
        #     listStockName.append(twstock.codes[code].name)
        # stockName = ",".join(listStockName)
        stockCode = sInStockCode
        stockName = twstock.codes[sInStockCode].name

        # 將查詢到的「股票名稱」帶入輸入框
        inputStockName.delete(0, tk.END)
        inputStockName.insert(0, stockName)

    if sInStockName != "":
        for code, info in twstock.codes.items():
            if sInStockName == info.name:
                stockCode = code
                stockName = info.name

                # 將查詢到的「股票名稱」帶入輸入框
                inputStockCode.delete(0, tk.END)
                inputStockCode.insert(0, stockCode)

    print(F"\t stockCode= {type(stockCode)} {stockCode}")
    print(F"\t stockName= {type(stockName)} {stockName}")


def getStockData():
    """
    依User輸入「股票代號/名稱」取得對應的技術指標相關資訊
    """
    print('===== getStockData() ===== ')
    global stockCode, stockName

    # ------------------------------
    # get User 輸入的 股票代號、股票名稱、開始日期、結束日期
    # ------------------------------
    # 先執行 showStockBasic 更新「股票代號」和「股票名稱」
    showStockBasic(inputStockCode, inputStockName)
    print(F"[getStockData] stockCode= {type(stockCode)} {stockCode}")
    print(F"[getStockData] stockName= {type(stockName)} {stockName}")

    # get 輸入的「開始日期」、「結束日期」
    sStartDate = inputStartDate.get()
    sEndDate   = inputEndDate.get()
    print(F"[getStockData] sStartDate = {type(sStartDate)} {sStartDate}")
    print(F"[getStockData] sEndDate   = {type(sEndDate)} {sEndDate}")

    try:
        dStartDate = datetime.strptime(sStartDate, '%Y-%m-%d')
        dEndDate   = datetime.strptime(sEndDate, '%Y-%m-%d')
    except ValueError:
        messagebox.showinfo("Error", "日期格式錯誤，請輸入 YYYY-MM-DD 格式")
        return

    if dStartDate >= dEndDate:
        messagebox.showinfo("Error", "「開始日期」必須早於「結束日期」")
        return

    # ------------------------------
    # 取得各技術指標
    # ------------------------------
    # objStockUtils = StockUtils()
    listStockInfo = objStockUtils.getStockAnalyze(stockCode,sEndDate,sStartDate)
    dictDays = objStockUtils.countIncreaseDaysHistory(listStockInfo)
    print(F"[getStockData] listStockInfo = ({type(listStockInfo)}) (size ={len(listStockInfo)})")
    print(F"[getStockData] dictDays = ({type(dictDays)}) (size ={len(dictDays)})")

    # ------------------------------
    # 宣告 list 後續存放要匯出的內容 <class 'list'>
    # ------------------------------
    listOutput = []

    # ------------------------------
    # 將 List 用日期降序排序
    # ------------------------------
    listStockInfo.sort(key=lambda x: x[0], reverse=False)
    for a in range(len(listStockInfo)):
        listStock = listStockInfo[a]
        stockDate   = listStock[0] # 交易日期
        stockCode   = listStock[1] # 股票代號
        stockName   = listStock[2] # 股票名稱
        fClosePrice = listStock[3] # 收盤價
        iUpCount    = listStock[4] # 多頭數量
        iDownCount  = listStock[5] # 空頭數量

        # get 「多頭/空頭數量」連續增加天數
        iUpDays   = dictDays[stockDate+"Up"]
        iDownDays = dictDays[stockDate+"Down"]
        print(F"\t [getStockData] {a}.{stockDate} {objStockUtils.sHr1}")
        print(f"\t iUpDays = {iUpDays}")
        print(f"\t iDownDays = {iDownDays}")

        listOutput.append([stockDate, stockCode, stockName, fClosePrice, iUpCount, iDownCount, iUpDays, iDownDays])
    print(F"[getStockData] listOutput = ({type(listOutput)}) (size ={len(listOutput)})")

    # ------------------------------
    # 將 結果 寫到「結果顯示文本框」
    # ------------------------------
    ilistOutput = len(listOutput)
    if ilistOutput > 0:
        header = f"股票代號: {stockCode}\n股票名稱: {stockName}\n日期範圍: {sStartDate} 至 {sEndDate}\n\n"
        header = f"{header} 第一筆的「連續多頭增加」及「連續空頭增加」因為沒有上一筆，故一定為0\n\n"
        print(F"[getStockData] header = ({type(header)}) {header}")

        # 將「結果顯示文本框」狀態設置為 NORMAL ，允許修改內容
        textResult.config(state=tk.NORMAL)
        # 刪除「結果顯示文本框」全部內容
        textResult.delete(1.0, tk.END)

        # 設置為微軟正黑體
        microsoft_zh_font = font.Font(family="Microsoft JhengHei", size=10)
        textResult.configure(font=microsoft_zh_font)

        textResult.insert(tk.END, header)
        textResult.insert(tk.END, '\n'.join([f"交易日期: {r[0]}, 收盤價: {r[3]:.2f}, 多頭指標: {r[4]}, 空頭指標: {r[5]}, 連續多頭增加: {r[6]}, 連續空頭增加: {r[7]}" for r in listOutput]))

        # 設置 sStartDate 和 sEndDate 顏色
        start_index = textResult.search(sStartDate, "1.0", tk.END)
        end_index   = textResult.search(sEndDate, "1.0", tk.END)
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
            print(F"[getStockData] sPath = ({type(sPath)}) {sPath}")

            objStockUtils.saveToExcel(sPath, listStockInfo, dictDays)
            messagebox.showinfo("Msg", "已查詢完成！並保存Excel")
        else:
            messagebox.showinfo("Msg", "已查詢完成！")



def readExcel():
    """
    依「演算法Excel」的「股票代號/名稱」取得對應的技術指標相關資訊
    """
    print('===== readExcel() ===== ')
    global stockCode, stockName
    
    # get User 選擇Excel的檔案路徑+檔名
    sFilePath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    # print(F"\t [readExcel] sFilePath = ({type(sFilePath)}) {sFilePath}")
    
    if not sFilePath:
        return

    try:
        # 將「結果顯示文本框」狀態設置為 NORMAL ，允許修改內容
        textResult.config(state=tk.NORMAL)
        # 刪除「結果顯示文本框」全部內容
        textResult.delete(1.0, tk.END)

        # 設置為微軟正黑體
        microsoft_zh_font = font.Font(family="Microsoft JhengHei", size=10)
        textResult.configure(font=microsoft_zh_font)

        # ------------------------------
        # 讀取 Excel 文件
        # ------------------------------
        df_excel = pd.read_excel(sFilePath)
        df_excel['日期'] = df_excel['日期'].astype(str)
        df_excel['股票代號'] = df_excel['股票代號'].astype(str)
        listStockCode = df_excel.iloc[:, 1].astype(str).tolist()
        print(F"[readExcel] listStockCode = ({type(listStockCode)}) (size ={len(listStockCode)}) {listStockCode}")
        
        # 創建了一個空的 DataFrame，用來存儲從多個股票代碼中獲取的所有股票數據
        df_result = pd.DataFrame()
        
        listOutput = []
        iNo = 0
        for stockCode in listStockCode:
            iNo += 1
            stockName = twstock.codes[stockCode].name
            print(F"[readExcel] {iNo}.({type(stockCode)}) {stockCode} {stockName} {objStockUtils.sHr1}")

            # ------------------------------
            # 取得各技術指標
            # ------------------------------
            listStockInfo = objStockUtils.getStockAnalyze(stockCode,"","")
            dictDays = objStockUtils.countIncreaseDaysHistory(listStockInfo)
            print(F"[readExcel] listStockInfo = ({type(listStockInfo)}) (size ={len(listStockInfo)})")
            print(F"[readExcel] dictDays = ({type(dictDays)}) (size ={len(dictDays)})")

            listStockInfo.sort(key=lambda x: x[0], reverse=False)
            for a in range(len(listStockInfo)):
                listStock = listStockInfo[a]
                stockDate   = listStock[0] # 交易日期
                stockCode   = listStock[1] # 股票代號
                stockName   = listStock[2] # 股票名稱
                fClosePrice = listStock[3] # 收盤價
                iUpCount    = listStock[4] # 多頭數量
                iDownCount  = listStock[5] # 空頭數量

                # get 「多頭/空頭數量」連續增加天數
                iUpDays   = dictDays[stockDate+"Up"]
                iDownDays = dictDays[stockDate+"Down"]
                # print(F"\t\t\t [readExcel] {a}.{stockDate} {objStockUtils.sHr1}")
                # print(F"\t\t\t [readExcel] stockDate = ({type(stockDate)}) ")
                # print(f"\treadExcel iUpDays = {iUpDays}")
                # print(f"\treadExcel iDownDays = {iDownDays}")
                
                listOutput.append([stockDate, stockCode, stockName, fClosePrice, iUpCount, iDownCount, iUpDays, iDownDays])

        print(F"[readExcel] listOutput = ({type(listOutput)}) (size ={len(listOutput)})")

        # 將資料輸出到GUI
        for a in range(len(listOutput)):
            listInfo = listOutput[a]
            print(F"[readExcel] listInfo =({type(listInfo)})  {listInfo}")
            stockDate   = listInfo[0] # 交易日期
            stockCode   = listInfo[1] # 股票代號
            stockName   = listInfo[2] # 股票名稱
            fClosePrice = listInfo[3] # 收盤價
            iUpCount    = listInfo[4] # 多頭 數量
            iDownCount  = listInfo[5] # 空頭 數量
            iUpDays     = listInfo[6] # 多頭 連續增加天數
            iDownDays   = listInfo[7] # 空頭 連續增加天數
            header = f"\n股票: {stockCode} {stockName}, 交易日期: {stockDate}, 收盤價: {fClosePrice}, 多頭指標: {iUpCount}, 空頭指標: {iDownCount}, 連續多頭增加: {iUpDays}, 連續空頭增加: {iUpDays}"
            textResult.insert(tk.END, header)

        # 將資料存到 DataFrame
        df_result = pd.DataFrame(listOutput, columns=["日期", "股票代號", "股票名稱", "收盤價", "技術指標多頭", "技術指標空頭", "連續多頭增加", "連續空頭增加"])


        #     # 將資料存到 DataFrame
        #     df_temp = pd.DataFrame(listOutput, columns=["日期", "股票代號", "股票名稱", "收盤價", "技術指標多頭", "技術指標空頭", "連續多頭增加", "連續空頭增加"])
        #     df_temp['股票代號'] = df_temp['股票代號'].astype(str)
            
        #     # 將每次迴圈中獲取的 df_temp 與 df_result 合併
        #     df_result = pd.concat([df_result, df_temp])

        # 使用 reset_index() 方法重置 df_result 的索引
        # drop=True 表示丟棄原來的索引，並創建一個新的默認整數索引
        # inplace=True 表示在原地修改 df_result，而不是返回一個新的 DataFrame
        df_result.reset_index(drop=True, inplace=True)

        df_result['日期'] = df_result['日期'].astype(str)
        df_result['股票代號'] = df_result['股票代號'].astype(str)

        # 將兩個 DataFrame 合併 on為合併的基準列
        # df_merged = pd.merge(df_excel, df_result, on=["日期", "股票代號", "股票名稱", "收盤價"], how='left')
        df_merged = pd.merge(df_excel, df_result, on=["日期", "股票代號", "股票名稱"], how='left')
        
        # 如果指定欄位存在，將其刪除
        listDelColumns = ['技術多頭', '技術空頭']
        isColumnsExist = [col for col in listDelColumns if col in df_merged.columns]
        if isColumnsExist:
            df_merged.drop(columns=listDelColumns, axis=1, inplace=True)

        # 排序
        df_merged = df_merged.sort_values(['連續多頭增加', '技術指標多頭'], ascending=[False, False])

        # get User輸入的 資料夾路徑
        sPath = inputFolderPath.get()
        sFileName = os.path.basename(sFilePath)
        print(F"\t [readExcel] sFileName = ({type(sFileName)}) {sFileName})")

        # 將資料保存到 Excel 文件
        # sFilePath = f"{sPath}/{sFileName}"
        sFilePath = f"{sPath}/test222.xlsx"
        df_merged.to_excel(sFilePath, index=False)
        messagebox.showinfo("保存成功", f"資料已保存至 {sFilePath}")


    except Exception as e:
        print(F"\t [ERROR] = {e}")
        messagebox.showerror("錯誤", f"錯誤: 讀取 Excel 文件失敗: {e}")




def verify():
    """
    回測程式 (row=4)
    """
    print('\t ===== verify() ===== ')



# print("==================================================")
root = tk.Tk() # 建立 tkinter 視窗物件
root.title("股票技術指標查詢") # 設定標題
root.resizable(False, False)  # 使窗口大小固定


# ------------------------------
# GUI-股票代號/名稱 (row=0)
# ------------------------------
ttk.Label(root, text="股票代號").grid(column=0, row=0, padx=10, pady=5, sticky=tk.W)
inputStockCode = ttk.Entry(root, width=25) #創建了一個輸入框
inputStockCode.insert(0, "3703")
inputStockCode.grid(column=1, row=0, padx=10, pady=5, sticky=tk.W, columnspan=2)

ttk.Label(root, text="股票名稱").grid(column=3, row=0, padx=10, pady=5, sticky=tk.W)
inputStockName = ttk.Entry(root, width=25) #創建了一個輸入框
inputStockName.insert(0, "")
inputStockName.grid(column=4, row=0, padx=10, pady=5, sticky=tk.W, columnspan=2)

# 綁定 Enter 鍵事件
inputStockCode.bind("<Return>", lambda event: showStockBasic(inputStockCode, inputStockName))
inputStockName.bind("<Return>", lambda event: showStockBasic(inputStockCode, inputStockName))
# 綁定離開文字框事件
inputStockCode.bind("<FocusOut>", lambda event: showStockBasic(inputStockCode, inputStockName))
inputStockName.bind("<FocusOut>", lambda event: showStockBasic(inputStockCode, inputStockName))


# ------------------------------
# GUI-開始日期 (row=1)
# ------------------------------
ttk.Label(root, text="開始日期").grid(column=0, row=1, padx=10, pady=5, sticky=tk.W)

# Set 預設日期
dStartDate = objStockUtils.getStartDate(datetime.today(),"")
startDate = tk.StringVar(value=dStartDate.strftime('%Y-%m-%d'))
# startDate = tk.StringVar(value="2024-06-26")
inputStartDate = ttk.Entry(root, textvariable=startDate, width=10)
inputStartDate.grid(column=1, row=1, padx=10, pady=5, sticky=tk.W)

buttonStartDate = ttk.Button(root, text="選擇日期", command=lambda: pick_date(startDate))
buttonStartDate.grid(column=2, row=1, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-結束日期 (row=1)
# ------------------------------
ttk.Label(root, text="結束日期").grid(column=3, row=1, padx=10, pady=5, sticky=tk.W)
# Set 預設日期
dEndDate = objStockUtils.adjustDate(datetime.today())
endDate = tk.StringVar(value=dEndDate.strftime('%Y-%m-%d'))
inputEndDate = ttk.Entry(root, textvariable=endDate, width=10)
inputEndDate.grid(column=4, row=1, padx=10, pady=5, sticky=tk.W)

buttonEndDate = ttk.Button(root, text="選擇日期", command=lambda: pick_date(endDate))
buttonEndDate.grid(column=5, row=1, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-開始日期 提醒訊息 (row=2)
# ------------------------------
ttk.Label(root, text="「開始日期」要小於「結束日期 - 31個交易日」，否則部份技術指標會有誤差或無值的狀況").grid(column=0, row=2, padx=10, pady=5, sticky=tk.W, columnspan=5)


# ------------------------------
# GUI-匯出Excel勾選框 (row=3)
# ------------------------------
exportExcel = tk.BooleanVar(value=True) # value=True 表示 預設勾選
checkExportExcel = ttk.Checkbutton(root, text="匯出Excel", variable=exportExcel)
checkExportExcel.grid(column=0, row=3, padx=10, pady=5, sticky=tk.W)

# 創建資料夾路徑輸入框
ttk.Label(root, text="資料夾路徑").grid(column=1, row=3, padx=10, pady=5)
# Set 預設路徑
defaultPath = tk.StringVar(value="C:/Users/user/Desktop/股票技術指標輸出")
inputFolderPath = ttk.Entry(root, textvariable=defaultPath, width=40)
inputFolderPath.grid(column=2, row=3, padx=10, pady=5, sticky=tk.W, columnspan=4)


# ------------------------------
# GUI-按鈕-查詢 (row=4)
# ------------------------------
buttonQuery = ttk.Button(root, text="查詢", command=getStockData)
buttonQuery.grid(column=0, row=4, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-按鈕-由演算法Excel取得股票代號 (row=4)
# ------------------------------
buttonExcel = ttk.Button(root, text="讀取Excel-演算法", command=readExcel)
buttonExcel.grid(column=1, row=4, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-按鈕-回測 (row=4)
# ------------------------------
buttonExcel = ttk.Button(root, text="讀取Excel-回測", command=verify)
buttonExcel.grid(column=2, row=4, padx=10, pady=5, sticky=tk.W)


# ------------------------------
# GUI-結果顯示
# ------------------------------
result_text = tk.StringVar()
# 加入 Frame 框架
#     sticky=(tk.W, tk.E, tk.N, tk.S)) : 使得框架可以隨窗口大小改變而伸展
frameResult = ttk.Frame(root)
frameResult.grid(column=0, row=5, padx=10, pady=5, sticky=(tk.W, tk.E, tk.N, tk.S), columnspan=6)

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
# 使用 mainloop() 將其放在主迴圈中一直執行，直到使用者關閉該視窗才會停止運作
# ------------------------------
root.mainloop()
