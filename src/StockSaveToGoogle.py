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


class StockSaveToGoogle:
    """
    取得「指定股票」的各項「技術指標」後，更新到 GoogleSheet.

    Attributes:
        sHr1 (str): 分隔線 in printed output.
        sHr2 (str): 分隔線 in printed output.

    Created on: 2023-06-23
    Last modified: 2024-06-23

    !!一直發生憑証錯誤
    An error occurred: ('invalid_grant: Invalid JWT Signature.', {'error': 'invalid_grant', 'error_description': 'Invalid JWT Signature.'})
    """

    # 排版相關變數
    sHr1 = "==================================================";
    sHr2 = "--------------------------------------------------";


    def __init__(self):
        """
        建構式: 建立初始化資料
        """
        print('===== class StockSaveToGoogle() ===== ')

        # 記錄 要「買」的股票
        self.listBuy = []

        # 記錄 要「賣」的股票
        self.listSell = []


        # new class
        self.objStockUtils = StockUtils()

        # ------------------------------
        # get 今天
        # ------------------------------
        self.dToday = datetime.today() #<class 'datetime.datetime'>
        # dToday_date_only = dToday.date()
        self.sToday = self.dToday.strftime('%Y-%m-%d') #<class 'str'>
        # get 星期 <class 'int'> (0 = Monday, 6 = Sunday)
        self.iWeekday = self.dToday.weekday() +1
        print(F"dToday = ({type(self.dToday)}) {self.dToday}")
        print(F"sToday = ({type(self.sToday)}) {self.sToday}")
        print(F"iWeekday = ({type(self.iWeekday)}) {self.iWeekday}")

        # ------------------------------
        # 使用憑證文件，開啟 GoogleSheet
        # ------------------------------
        # get 目前程式檔案路徑
        script_dir = os.path.dirname(__file__)

        # Set JSON 憑證文件的路徑
        credentials_file = os.path.join(script_dir, "stock-426606-5b695171527f.json")
        print(F"credentials_file = {type(credentials_file)} {credentials_file}")

        # 使用 JSON 憑證文件來授權
        self.gc = pygsheets.authorize(service_file=credentials_file)

        # 開啟 GoogleSheet 試算表(spreadsheet)
        self.shtInputBuy  = self.gc.open_by_key('1a6Z1LSBV2u7P01X1vT0v5AyqqwcH-aj_dQ9pbfXV1X4') # (要買的) 聰全資料表-選股記錄-成交量排行-20000
        self.shtInputSell = self.gc.open_by_key('1YjPSxs2gM0HNC095Y2JW7nbG83rPPWA4MAmvdNnmUMU') # (要賣的) 觀察股票List
        self.shtOutput = self.gc.open_by_key('18kL5aIbNLudSUkgfu1pdGVYf-2citOQjjCST3lIInEA') # 選股記錄_2024
        self.shtOutputDetails = self.gc.open_by_key('13ya5frtagmkCojhthWbG82L0NzTshX5PqwK25MNER3k') #技術指標明細


    def upCurrentStock(self):
        """
        取得「目前買的股票」的基本資料 及 各項技術指標，更新到 GoogleSheet
        """
        print('===== upCurrentStock() ===== ')

        try:
            # 組出要寫到 GoogleSheet 的內容
            listOutput = []

            # 處理每個工作表的數據
            sGetDay = self.sToday
            # sGetDay = "2024-07-12"

            # get GoogleSheet 中的第一個 工作表(worksheet)
            wksInput = self.shtInputSell[0]

            # 讀取「觀察股票List」，存到 DataFrame
            dfInput = pd.DataFrame(wksInput.get_all_records())
            # print(f"dfInput = {dfInput}")

            for row in dfInput.itertuples(index=True, name='Pandas'):
                i = row.Index
                stockCode = str(row[1])
                stockName = row[2]
                print(f"{self.sHr1}")
                print(f"[upCurrentStock] indexRow = {i}, 股票代號 = {stockCode}, 股票名稱 = {stockName}")

                # !!方便測試時，提早跳出迴圈
                # if i == 1:
                #     break

                # 防呆: 檢查 股票代號 是否為空值
                if pd.isna(stockCode) or stockCode == "":
                    continue  # 跳出此次迴圈，進入下一個迴圈

                # 爬Yahoo，取得當天股票資訊
                dictYahooData = self.objStockUtils.getStockData4Yahoo(stockCode)
                print(f"dictYahooData = {dictYahooData}")


                # 計算技術指標
                listStockInfo = self.objStockUtils.getStockAnalyze(stockCode,"","")
                dictDays = self.objStockUtils.countIncreaseDaysHistory(listStockInfo)
                print(F"[upCurrentStock] listStockInfo size = {type(listStockInfo)} {len(listStockInfo)}")

                for a in range(len(listStockInfo)):
                    listStock = listStockInfo[a]
                    stockDate   = listStock[0] # 資料日期
                    stockCode   = listStock[1] # 股票代號
                    stockName   = listStock[2] # 股票名稱
                    fClosePrice = listStock[3] # 收盤價
                    iUpQty      = listStock[4] # 多頭數量
                    iDownQty    = listStock[5] # 空頭數量
                    dictSignals = listStock[6] # 詳細各計術指標值、歷史價格數據
                    iUpDays00   = dictDays[stockDate+"Up"]
                    iDownDays00 = dictDays[stockDate+"Down"]

                    # 只取當天的資料，故日期非當天即跳出迴圈
                    if (stockDate != sGetDay):
                        break

                    print(F"\t [upCurrentStock] a = {a} {self.sHr2}")
                    print(F"\t stockDate = ({type(stockDate)}) {stockDate}")

                    # ------------------------------
                    # get 多頭指標
                    # ------------------------------
                    # 取得「多頭數量」連續 >=index 的天數
                    iUpDays20 = self.objStockUtils.countIndexDays('up', listStockInfo, 20)
                    iUpDays15 = self.objStockUtils.countIndexDays('up', listStockInfo, 15)
                    print(F"\t iUpDays00 = ({type(iUpDays00)}) {iUpDays00}")
                    print(F"\t iUpDays20 = ({type(iUpDays20)}) {iUpDays20}")
                    print(F"\t iUpDays15 = ({type(iUpDays15)}) {iUpDays15}")

                    # ------------------------------
                    # get 空頭指標
                    # ------------------------------
                    # 取得「空頭數量」連續 >=index 的天數
                    iDownDays20 = self.objStockUtils.countIndexDays('down', listStockInfo, 20)
                    iDownDays15 = self.objStockUtils.countIndexDays('down', listStockInfo, 15)
                    print(F"\t iDownDays00 = ({type(iDownDays00)}) {iDownDays00}")
                    print(F"\t iDownDays20 = ({type(iDownDays20)}) {iDownDays20}")
                    print(F"\t iDownDays15 = ({type(iDownDays15)}) {iDownDays15}")

                    # 將數據保存到 Excel 文件
                    sPath = "C:/Users/user/Desktop/股票技術指標輸出/"
                    self.objStockUtils.saveToExcel(sPath, listStockInfo, dictDays)

                    # 將數據保存到 Google 文件 (技術指標明細)
                    # self.objStockUtils.saveToGoogle(listStockInfo, dictDays)

                    # ------------------------------
                    # get 賣出Flag
                    # ------------------------------
                    # 「空頭數量」連續增加天數 >= 2 OR 「空頭數量」連續大於20天數 > 3
                    isSell = False
                    if iDownDays00 >= 2 or iDownDays20 > 3:
                        isSell = True

                    # 將新行添加到 dfOutput
                    listOutput.append({
                        "更新日期": stockDate,
                        "星期": self.iWeekday,
                        "股票代號": stockCode,
                        "股票名稱": stockName,
                        "賣出Flag": isSell,
                        "最新股價": "",
                        "最新報酬率" : "",
                        "收盤價": fClosePrice,
                        "漲跌": dictYahooData.get("漲跌", 0),
                        "幅度(%)": dictYahooData.get("幅度", 0),
                        "成交量": dictYahooData.get("成交量", 0),
                        "連漲連跌": dictYahooData.get("連漲連跌", 0),
                        "技術指標-多頭數量" : iUpQty,
                        "技術指標-空頭數量" : iDownQty,
                        "「多頭數量」連續增加天數" : iUpDays00,
                        "「多頭數量」連續>20天數"  : iUpDays20,
                        "「多頭數量」連續>15天數"  : iUpDays15,
                        "「空頭數量」連續增加天數" : iDownDays00,
                        "「空頭數量」連續>20天數"  : iDownDays20,
                        "「空頭數量」連續>15天數"  : iDownDays15,
                        "簡單移動平均線 (SMA)" : dictSignals.get("簡單移動平均線 (SMA)", 0),
                        "指數移動平均線 (EMA)" : dictSignals.get("指數移動平均線 (EMA)", 0),
                        "移動平均收斂背離指標 (MACD)" : dictSignals.get("移動平均收斂背離指標 (MACD)", 0),
                        "相對強弱指數 (RSI)" : dictSignals.get("相對強弱指數 (RSI)", 0),
                        "平衡交易量 (OBV)" : dictSignals.get("平衡交易量 (OBV)", 0),
                        "威廉指標 (WILLR)" : dictSignals.get("威廉指標 (WILLR)", 0),
                        "動量指標 (MOM)" : dictSignals.get("動量指標 (MOM)", 0),
                        "龐氏指標 (SAR)" : dictSignals.get("龐氏指標 (SAR)", 0),
                        "力量指標 (FORCE)" : dictSignals.get("力量指標 (FORCE)", 0),
                        "成交量比率 (VROC)" : dictSignals.get("成交量比率 (VROC)", 0),
                        "Money Flow Index (MFI)" : dictSignals.get("Money Flow Index (MFI)", 0),
                        # add by huey,2024/07/07,加入新的技術指標
                        "線性回歸 (LINEARREG)" : dictSignals.get("線性回歸 (LINEARREG)", 0),
                        "線性回歸角度 (LINEARREG_ANGLE)" : dictSignals.get("線性回歸角度 (LINEARREG_ANGLE)", 0),
                        "線性回歸斜率 (LINEARREG_SLOPE)" : dictSignals.get("線性回歸斜率 (LINEARREG_SLOPE)", 0),
                        "中價 (MEDPRICE)" : dictSignals.get("中價 (MEDPRICE)", 0),
                        "Percentage Price Oscillator (PPO)" : dictSignals.get("Percentage Price Oscillator (PPO)", 0),
                        "Rate of change (ROC)" : dictSignals.get("Rate of change (ROC)", 0),
                        "Rate of change Percentage (ROCP)" : dictSignals.get("Rate of change Percentage (ROCP)", 0),
                        "時間序列預測 (TSF)" : dictSignals.get("時間序列預測 (TSF)", 0),
                        "典型價格 (TYPPRICE)" : dictSignals.get("典型價格 (TYPPRICE)", 0),
                        "加權收盤價格 (WCLPRICE)" : dictSignals.get("加權收盤價格 (WCLPRICE)", 0),
                        # add by huey,2024/07/07,移除的技術指標
                        "隨機指標 (STOCH)" : dictSignals.get("隨機指標 (STOCH)", 0),
                        "布林帶 (BBANDS)" : dictSignals.get("布林帶 (BBANDS)", 0),
                        "平均真實範圍 (ATR)" : dictSignals.get("平均真實範圍 (ATR)", 0),
                        "商品通道指數 (CCI)" : dictSignals.get("商品通道指數 (CCI)", 0),
                        "平均方向性指數 (ADX)" : dictSignals.get("平均方向性指數 (ADX)", 0),
                        "可變移動平均線 (VARMA)" : dictSignals.get("可變移動平均線 (VARMA)", 0),
                        "成交量移動平均線 (VMA 短期)" : dictSignals.get("成交量移動平均線 (VMA 短期)", 0),
                        "成交量移動平均線 (VMA 中期)" : dictSignals.get("成交量移動平均線 (VMA 中期)", 0),
                        "成交量移動平均線 (VMA 長期)" : dictSignals.get("成交量移動平均線 (VMA 長期)", 0),
                        "Chande Momentum Oscillator (CMO)" : dictSignals.get("Chande Momentum Oscillator (CMO)", 0),
                        "Kaufman Adaptive Moving Average (KAMA)" : dictSignals.get("Kaufman Adaptive Moving Average (KAMA)", 0),
                    })

            # ------------------------------
            # 將資料寫入 GoogleSheet 的內容
            # ------------------------------
            print(f"[upBuyingStock] 開始更新 GoogleSheet")
            # Set 表頭
            columns = ["更新日期","星期","股票代號","股票名稱"
                        ,"賣出Flag","最新股價","最新報酬率","收盤價","漲跌","幅度(%)","成交量","連漲連跌"
                        ,"技術指標-多頭數量","技術指標-空頭數量"
                        ,"「多頭數量」連續增加天數","「多頭數量」連續>20天數","「多頭數量」連續>15天數"
                        ,"「空頭數量」連續增加天數","「空頭數量」連續>20天數","「空頭數量」連續>15天數"
                        , "簡單移動平均線 (SMA)","指數移動平均線 (EMA)","移動平均收斂背離指標 (MACD)","相對強弱指數 (RSI)","平衡交易量 (OBV)","威廉指標 (WILLR)","動量指標 (MOM)","龐氏指標 (SAR)","力量指標 (FORCE)","成交量比率 (VROC)","Money Flow Index (MFI)","線性回歸 (LINEARREG)","線性回歸角度 (LINEARREG_ANGLE)","線性回歸斜率 (LINEARREG_SLOPE)","中價 (MEDPRICE)","Percentage Price Oscillator (PPO)","Rate of change (ROC)","Rate of change Percentage (ROCP)","時間序列預測 (TSF)","典型價格 (TYPPRICE)","加權收盤價格 (WCLPRICE)",
                        "隨機指標 (STOCH)","布林帶 (BBANDS)","平均真實範圍 (ATR)","商品通道指數 (CCI)","平均方向性指數 (ADX)","可變移動平均線 (VARMA)","成交量移動平均線 (VMA 短期)","成交量移動平均線 (VMA 中期)","成交量移動平均線 (VMA 長期)","Chande Momentum Oscillator (CMO)","Kaufman Adaptive Moving Average (KAMA)"
                        ]
            print(F"\t columns size {type(columns)} {len(columns)}")
            print(F"\t listOutput size {type(listOutput)} {len(listOutput)}")

            dfOutput = pd.DataFrame(listOutput, columns=columns)

            # 將資料依「技術指標-多頭數量」排序
            dfOutput = dfOutput.sort_values(['「多頭數量」連續增加天數', '技術指標-多頭數量'], ascending=[False, False])

            # get 目前工作表資料，並轉為 DataFrame
            wksOutput = self.shtOutput[1]
            existingData = wksOutput.get_all_records()
            dfExisting = pd.DataFrame(existingData)

            # 將新的資料附加到現有資料後面
            dfCombined = pd.concat([dfExisting, dfOutput], ignore_index=True)

            # 將更新後的 DataFrame 寫回工作表
            wksOutput.set_dataframe(dfCombined, (0, 0))
            print(F"Google 已更新. 更新筆數:{len(dfOutput)}")


        except Exception as e:
            print(f"An error occurred: {e}")



    def processWorksheet(self, worksheet, sOutputDate):
        '''
        取得工作表，整理後加到「每日演算法選出要買的股票」清單中

        Args:
            worksheet (pygsheets.worksheet.Worksheet): GoogleSheet 工作表物件.
        Returns:
            沒有回傳任何東西
        '''
        # print('\t ===== processWorksheet() ===== ')
        # print(F"\t worksheet = {type(worksheet)}")
        # print(F"\t sOutputDate = {type(sOutputDate)} {sOutputDate}")

        # get 工作表數據
        rows = worksheet.get_all_values()
        # print(F"\t rows = {type(rows)} {len(rows)}")
        # 讀取「工作表(worksheet)」，存到 DataFrame
        # dfInput1 = pd.DataFrame(worksheet.get_all_records())
        
        # get 表頭
        headers = rows[0]
        # print(F"\t headers = {type(headers)} {headers}")

        # get 頁籤名稱
        data_source = worksheet.title

        i = 0
        for row in rows[1:]:
        # for row in dfInput1.itertuples(index=True, name='Pandas'):
            i += 1
            stockDate = row[0]
            stockCode = row[1]
            stockName = row[2]

            # 只抓指定日期的資料
            if sOutputDate != stockDate:
                continue

            # 建立字典，headers[i] 為 key，row[i] 為 value
            # 從索引值 3 開始到 headers 列表的最後一個元素的索引（len(headers) - 1）才加入
            additional_data = {
                headers[i]: row[i] for i in range(3, len(headers))
            }
            # print(f"\t {self.sHr1}")
            # print(f"\t indexRow = {i}, 資料日期 = {stockDate}, 股票代號 = {stockCode}, 股票名稱 = {stockName}")
            # print(F"\t additional_data = {type(additional_data)} {additional_data}")
            # print(F"\t data_source = {type(data_source)} {data_source}")

            # 檢查「日期+股票代號」是否已有存在於列表中，
            # 如果存在，更新資料; 否則，新增資料
            found = False
            for item in self.listBuy:
                if item[0] == stockDate and item[1] == stockCode:
                    item[4].add(data_source)
                    found = True
                    break

            if not found:
                self.listBuy.append([stockDate, stockCode, stockName, additional_data, {data_source}])


    def upBuyingStock(self):
        """
        取得「預計要買的股票」的基本資料 及 各項技術指標，更新到 GoogleSheet
        """
        print('===== upBuyingStock() ===== ')

        try:
            # 組出要寫到 GoogleSheet 的內容
            listOutput = []

            # ------------------------------
            # 讀取工作表-選股記錄_2024，整理後存到 self.listBuy
            # ------------------------------
            # wksOutput = self.shtOutput.worksheet_by_title('預計要「買」的')
            # # 取得工作表中所有資料
            # all_data = wksOutput.get_all_values()
            # # 計算資料筆數
            # num_rows = 0
            # for row in all_data[1:]:
            #     num_rows += 1
            #     stockDate = row[0]
            #     if stockDate == "":
            #         break
            # print(F"num_rows = {type(num_rows)} {num_rows}")



            # ------------------------------
            # 讀取工作表-聰全資料表，整理後存到 self.listBuy
            # ------------------------------
            # get GoogleSheet 中的工作表(worksheet)
            wksInput1 = self.shtInputBuy.worksheet_by_title('支持向量機-20000筆') #支持向量機
            wksInput2 = self.shtInputBuy.worksheet_by_title('梯度提升樹-20000筆') #梯度提升樹
            wksInput3 = self.shtInputBuy.worksheet_by_title('MPL-20000筆') #MPL-20000筆

            # 處理每個工作表的數據
            sGetDay = self.sToday
            # sGetDay = "2024-07-12"
            self.processWorksheet(wksInput1,sGetDay)
            self.processWorksheet(wksInput2,sGetDay)
            self.processWorksheet(wksInput3,sGetDay)
            print(F"self.listBuy size = {type(self.listBuy)} {len(self.listBuy)}")


            # 印出 list 中的資料
            # for item in self.listBuy:
            #     print(f"\t {self.sHr1}")
            #     print(f"\t 日期: {item[0]}, 股票代號: {item[1]}, 股票名稱: {item[2]}")
            #     print(f"\t 詳細資料: {item[3]}")
            #     print(f"\t 資料來源: {item[4]}")


            # ------------------------------
            # 組出要寫到 GoogleSheet 的內容
            # ------------------------------
            iRow = 0
            for item in self.listBuy:
                iRow += 1
                stockDate = item[0] # 資料日期
                stockCode = item[1] # 股票代號
                stockName = item[2] # 股票名稱
                dict      = item[3] # 詳細數據
                dataSource= item[4] # 資料來源
                print(f"{self.sHr1}")
                print(f"[upBuyingStock] indexRow = {iRow}, 資料日期 = {stockDate}, 股票代號 = {stockCode}, 股票名稱 = {stockName}")

                # !!方便測試時，提早跳出迴圈
                # if iRow == 2:
                #     break

                # 取得各技術指標
                result = []
                listStockInfo = self.objStockUtils.getStockAnalyze(stockCode,"","")
                dictDays = self.objStockUtils.countIncreaseDaysHistory(listStockInfo)
                print(F"[upBuyingStock] listStockInfo size = {type(listStockInfo)} {len(listStockInfo)}")

                for a in range(len(listStockInfo)):
                    listStock = listStockInfo[a]
                    stockDate   = listStock[0] # 資料日期
                    stockCode   = listStock[1] # 股票代號
                    stockName   = listStock[2] # 股票名稱
                    fClosePrice = listStock[3] # 收盤價
                    iUpQty      = listStock[4] # 多頭數量
                    iDownQty    = listStock[5] # 空頭數量
                    dictSignals = listStock[6] # 詳細各計術指標值、歷史價格數據
                    iUpDays00   = dictDays[stockDate+"Up"]    # 「空頭數量」連續增加天數
                    iDownDays00 = dictDays[stockDate+"Down"]  # 「多頭數量」連續增加天數

                    # 只取當天的資料，故日期非當天即跳出迴圈
                    # if (stockDate != "2024-06-25"):
                    #     continue
                    # if (stockDate != self.sToday):
                    if (stockDate != sGetDay):
                        break

                    print(F"\t [upBuyingStock] a = {a} {self.sHr2}")


                    # ------------------------------
                    # get 多頭指標
                    # ------------------------------
                    # 取得「多頭數量」連續 >=index 的天數
                    iUpDays20 = self.objStockUtils.countIndexDays('up', listStockInfo, 20)
                    iUpDays15 = self.objStockUtils.countIndexDays('up', listStockInfo, 15)
                    print(F"\t iUpDays00 = ({type(iUpDays00)}) {iUpDays00}")
                    print(F"\t iUpDays20 = ({type(iUpDays20)}) {iUpDays20}")
                    print(F"\t iUpDays15 = ({type(iUpDays15)}) {iUpDays15}")

                    # ------------------------------
                    # get 空頭指標
                    # ------------------------------
                    # 取得「空頭數量」連續 >=index 的天數
                    iDownDays20 = self.objStockUtils.countIndexDays('down', listStockInfo, 20)
                    iDownDays15 = self.objStockUtils.countIndexDays('down', listStockInfo, 15)
                    print(F"\t iDownDays00 = ({type(iDownDays00)}) {iDownDays00}")
                    print(F"\t iDownDays20 = ({type(iDownDays20)}) {iDownDays20}")
                    print(F"\t iDownDays15 = ({type(iDownDays15)}) {iDownDays15}")

                    # 將數據保存到 Excel 文件
                    sPath = "C:/Users/user/Desktop/股票技術指標輸出"
                    self.objStockUtils.saveToExcel(sPath, listStockInfo, dictDays)

                    # 將數據保存到 Google 文件 (技術指標明細)
                    self.objStockUtils.saveToGoogle(listStockInfo, dictDays)

                    # ------------------------------
                    # get 買進Flag
                    # ------------------------------
                    # 5日漲幅 < 5
                    # 20日漲幅 < 10
                    # 外資連買(天) > 1 OR 投信連買(天) > 1
                    # 進出分點總家數差 < 0
                    # 「多頭數量」連續增加天數 >= 2
                    growth05 = dict.get("5日漲幅", 0)
                    growth20 = dict.get("20日漲幅", 0)
                    foreign_investors_buying = dict.get("外資連買(天)", 0)
                    investment_trust_buying  = dict.get("投信連買(天)", 0)
                    households = dict.get("進出分點總家數差", 0)
                    print(F"\t 5日漲幅 = ({type(growth05)}) {growth05}")
                    print(F"\t 20日漲幅 = ({type(growth20)}) {growth20}")
                    print(F"\t 外資連買(天) = ({type(foreign_investors_buying)}) {foreign_investors_buying}")
                    print(F"\t 投信連買(天) = ({type(investment_trust_buying)}) {investment_trust_buying}")
                    print(F"\t 進出分點總家數差 = ({type(households)}) {households}")

                    isBuy = False
                    if float(growth05) < 5 and float(growth20) < 10:
                        if (float(foreign_investors_buying) > 1 or float(investment_trust_buying) > 1):
                            if int(households) < 0:
                                if iUpDays00 >= 2:
                                    isBuy = True

                    listOutput.append({
                        "更新日期": stockDate,
                        "星期": self.iWeekday,
                        "股票代號": stockCode,
                        "股票名稱": stockName,
                        "買進Flag": isBuy,
                        "最新股價": "",#f"=GOOGLEFINANCE(\"TPE:\"&$C{num_rows+1+iRow}, \"price\")",
                        "最新報酬率": "",#f"=($F{num_rows+1+iRow}-$H{num_rows+1+iRow})/$H{num_rows+1+iRow}",
                        "收盤價": fClosePrice,
                        "漲幅(%)": dict.get("漲幅(%)", 0),
                        "技術指標-多頭數量" : iUpQty,
                        "技術指標-空頭數量" : iDownQty,
                        "「多頭數量」連續增加天數" : iUpDays00,
                        "「多頭數量」連續>20天數"  : iUpDays20,
                        "「多頭數量」連續>15天數"  : iUpDays15,
                        "「空頭數量」連續增加天數" : iDownDays00,
                        "「空頭數量」連續>20天數"  : iDownDays20,
                        "「空頭數量」連續>15天數"  : iDownDays15,
                        "5日漲幅" : dict.get("5日漲幅", 0),
                        "20日漲幅": dict.get("20日漲幅", 0),
                        "外資連買(天)": dict.get("外資連買(天)", 0),
                        "外資連買張數": dict.get("外資連買張數", 0),
                        "投信連買(天)": dict.get("投信連買(天)", 0),
                        "投信連買張數": dict.get("投信連買張數", 0),
                        "自營商連買(天)": dict.get("自營商連買(天)", 0),
                        "大戶近1週增減"  : dict.get("大戶近1週增減", 0),
                        "散戶近1週增減％": dict.get("散戶近1週增減％", 0),
                        "進出分點總家數差": dict.get("進出分點總家數差", 0),
                        "產業名稱": dict.get("產業名稱", 0),
                        "資料來源" : dataSource,
                        "簡單移動平均線 (SMA)" : dictSignals.get("簡單移動平均線 (SMA)", 0),
                        "指數移動平均線 (EMA)" : dictSignals.get("指數移動平均線 (EMA)", 0),
                        "移動平均收斂背離指標 (MACD)" : dictSignals.get("移動平均收斂背離指標 (MACD)", 0),
                        "相對強弱指數 (RSI)" : dictSignals.get("相對強弱指數 (RSI)", 0),
                        "平衡交易量 (OBV)" : dictSignals.get("平衡交易量 (OBV)", 0),
                        "威廉指標 (WILLR)" : dictSignals.get("威廉指標 (WILLR)", 0),
                        "動量指標 (MOM)" : dictSignals.get("動量指標 (MOM)", 0),
                        "龐氏指標 (SAR)" : dictSignals.get("龐氏指標 (SAR)", 0),
                        "力量指標 (FORCE)" : dictSignals.get("力量指標 (FORCE)", 0),
                        "成交量比率 (VROC)" : dictSignals.get("成交量比率 (VROC)", 0),
                        "Money Flow Index (MFI)" : dictSignals.get("Money Flow Index (MFI)", 0),
                        # add by huey,2024/07/07,加入新的技術指標
                        "線性回歸 (LINEARREG)" : dictSignals.get("線性回歸 (LINEARREG)", 0),
                        "線性回歸角度 (LINEARREG_ANGLE)" : dictSignals.get("線性回歸角度 (LINEARREG_ANGLE)", 0),
                        "線性回歸斜率 (LINEARREG_SLOPE)" : dictSignals.get("線性回歸斜率 (LINEARREG_SLOPE)", 0),
                        "中價 (MEDPRICE)" : dictSignals.get("中價 (MEDPRICE)", 0),
                        "Percentage Price Oscillator (PPO)" : dictSignals.get("Percentage Price Oscillator (PPO)", 0),
                        "Rate of change (ROC)" : dictSignals.get("Rate of change (ROC)", 0),
                        "Rate of change Percentage (ROCP)" : dictSignals.get("Rate of change Percentage (ROCP)", 0),
                        "時間序列預測 (TSF)" : dictSignals.get("時間序列預測 (TSF)", 0),
                        "典型價格 (TYPPRICE)" : dictSignals.get("典型價格 (TYPPRICE)", 0),
                        "加權收盤價格 (WCLPRICE)" : dictSignals.get("加權收盤價格 (WCLPRICE)", 0),
                        # add by huey,2024/07/07,移除的技術指標
                        "隨機指標 (STOCH)" : dictSignals.get("隨機指標 (STOCH)", 0),
                        "布林帶 (BBANDS)" : dictSignals.get("布林帶 (BBANDS)", 0),
                        "平均真實範圍 (ATR)" : dictSignals.get("平均真實範圍 (ATR)", 0),
                        "商品通道指數 (CCI)" : dictSignals.get("商品通道指數 (CCI)", 0),
                        "平均方向性指數 (ADX)" : dictSignals.get("平均方向性指數 (ADX)", 0),
                        "可變移動平均線 (VARMA)" : dictSignals.get("可變移動平均線 (VARMA)", 0),
                        "成交量移動平均線 (VMA 短期)" : dictSignals.get("成交量移動平均線 (VMA 短期)", 0),
                        "成交量移動平均線 (VMA 中期)" : dictSignals.get("成交量移動平均線 (VMA 中期)", 0),
                        "成交量移動平均線 (VMA 長期)" : dictSignals.get("成交量移動平均線 (VMA 長期)", 0),
                        "Chande Momentum Oscillator (CMO)" : dictSignals.get("Chande Momentum Oscillator (CMO)", 0),
                        "Kaufman Adaptive Moving Average (KAMA)" : dictSignals.get("Kaufman Adaptive Moving Average (KAMA)", 0),
                    })
                    print(F"a = {a} |listOutput size {type(listOutput)} {len(listOutput)}")

            # ------------------------------
            # 將資料寫入 GoogleSheet 的內容
            # ------------------------------
            print(f"[upBuyingStock] 開始更新 GoogleSheet")
            # Set 表頭
            columns = ["更新日期","星期","股票代號","股票名稱"
                        ,"買進Flag","最新股價","最新報酬率","收盤價","漲幅(%)"
                        ,"技術指標-多頭數量","技術指標-空頭數量"
                        ,"「多頭數量」連續增加天數","「多頭數量」連續>20天數","「多頭數量」連續>15天數"
                        ,"「空頭數量」連續增加天數","「空頭數量」連續>20天數","「空頭數量」連續>15天數"
                        ,"5日漲幅","20日漲幅"
                        ,"外資連買(天)","外資連買張數","投信連買(天)","投信連買張數","自營商連買(天)"
                        ,"大戶近1週增減","散戶近1週增減％","進出分點總家數差","產業名稱","資料來源"
                        , "簡單移動平均線 (SMA)","指數移動平均線 (EMA)","移動平均收斂背離指標 (MACD)","相對強弱指數 (RSI)","平衡交易量 (OBV)","威廉指標 (WILLR)","動量指標 (MOM)","龐氏指標 (SAR)","力量指標 (FORCE)","成交量比率 (VROC)","Money Flow Index (MFI)","線性回歸 (LINEARREG)","線性回歸角度 (LINEARREG_ANGLE)","線性回歸斜率 (LINEARREG_SLOPE)","中價 (MEDPRICE)","Percentage Price Oscillator (PPO)","Rate of change (ROC)","Rate of change Percentage (ROCP)","時間序列預測 (TSF)","典型價格 (TYPPRICE)","加權收盤價格 (WCLPRICE)",
                        "隨機指標 (STOCH)","布林帶 (BBANDS)","平均真實範圍 (ATR)","商品通道指數 (CCI)","平均方向性指數 (ADX)","可變移動平均線 (VARMA)","成交量移動平均線 (VMA 短期)","成交量移動平均線 (VMA 中期)","成交量移動平均線 (VMA 長期)","Chande Momentum Oscillator (CMO)","Kaufman Adaptive Moving Average (KAMA)"
                        ]
            print(F"columns size {type(columns)} {len(columns)}")
            print(F"listOutput size {type(listOutput)} {len(listOutput)}")

            dfOutput = pd.DataFrame(listOutput, columns=columns)

            # 將資料依「技術指標-多頭數量」排序
            # dfOutput = dfOutput.sort_values(by='技術指標-多頭數量', ascending=False)
            dfOutput = dfOutput.sort_values(['「多頭數量」連續增加天數', '技術指標-多頭數量'], ascending=[False, False])

            # get 目前工作表資料，並轉為 DataFrame
            wksOutput = self.shtOutput[0]
            existingData = wksOutput.get_all_records()
            dfExisting = pd.DataFrame(existingData)

            # 將新的資料附加到現有資料後面
            dfCombined = pd.concat([dfExisting, dfOutput], ignore_index=True)

            # 將更新後的 DataFrame 寫回工作表
            wksOutput.set_dataframe(dfCombined, (0, 0))
            print(F"Google 已更新. 更新筆數:{len(dfOutput)}")

            # 這個方式程式效能很差，要再測試，暫註解
            # 設置公式
            # formulas = [
            #     ('F', '=GOOGLEFINANCE("TPE:"&$C{row}, "price")'),  # 最新股價
            #     ('G', '=($F{row}-$H{row})/$H{row}')  # 最新報酬率
            # ]

            # # for cell, formula in formulas:
            # #     wksOutput.update_value(cell, formula)

            # 遍歷所有行，並將公式寫入
            # start_row = 2
            # end_row = 200
            # for row in range(start_row, end_row + 1):
            #     for col, formula in formulas:
            #         cell = f'{col}{row}'
            #         formatted_formula = formula.format(row=row)
            #         wksOutput.update_value(cell, formatted_formula)

        except Exception as e:
            print(f"[ERROR]upBuyingStock() in StockSaveToGoogle: {e}")




# ==================================================
# new class
obj = StockSaveToGoogle()

# obj.upCurrentStock()

obj.upBuyingStock()