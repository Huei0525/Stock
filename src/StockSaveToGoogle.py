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
    """

    # 排版相關變數
    sHr1 = "==================================================";
    sHr2 = "--------------------------------------------------";


    def __init__(self):
        """
        建構式: 建立初始化資料
        """
        print('\t ===== class StockSaveToGoogle() ===== ')

        # get 目前程式檔案路徑
        script_dir = os.path.dirname(__file__)

        # Set JSON 憑證文件的路徑
        credentials_file = os.path.join(script_dir, "stock-426606-d2563f95c1b5.json")
        print(F"\t credentials_file = {type(credentials_file)} {credentials_file}")

        # 使用 JSON 憑證文件來授權
        self.gc = pygsheets.authorize(service_file=credentials_file)


    def upCurrentStock(self):
        """
        取得「目前買的股票」的基本資料 及 各項技術指標，更新到 GoogleSheet
        """
        print('\t ===== upCurrentStock() ===== ')

        try:
            # ------------------------------
            # 前置作業，使用憑證文件，開啟 GoogleSheet
            # ------------------------------
            # 開啟 GoogleSheet 試算表(spreadsheet)
            # shtTest   = gc.open_by_key('1djPZNb2FfyLBLNWeuTwA9TNWLzdCaR5IRTOjy1Zn36w') # 測試試算表
            shtInput  = self.gc.open_by_key('1YjPSxs2gM0HNC095Y2JW7nbG83rPPWA4MAmvdNnmUMU') # 觀察股票List
            shtOutput = self.gc.open_by_key('1NvtyIQ0370gUCngRgaTUNxPs7kRTBpT1tSnsjl0IfMY') # 技術指標List


            # get GoogleSheet 中的第一個 工作表(worksheet)
            wksInput  = shtInput[0]
            wksOutput = shtOutput[0]

            # 讀取「觀察股票List」，存到 DataFrame
            dfInput = pd.DataFrame(wksInput.get_all_records())
            # print(f"dfInput = {dfInput}")

            # Create empty List，用於收進要寫入「技術指標List」的資料
            listOutput = []

            # ------------------------------
            # 讀取「觀察股票List」
            # ------------------------------
            # get 當天日期 <class 'str'>
            sToday = datetime.today().strftime('%Y-%m-%d')
            # get 星期 <class 'int'> (0 = Monday, 6 = Sunday)
            iWeekday = datetime.strptime(sToday, '%Y-%m-%d').weekday() +1
            print(F"sToday = {sToday} ,iWeekday = {iWeekday}")


        except Exception as e:
            print(f"An error occurred: {e}")



    def upBuyingStock(self):
        """
        取得「預計要買的股票」的基本資料 及 各項技術指標，更新到 GoogleSheet
        """
        print('\t ===== upBuyingStock() ===== ')

        try:
            # ------------------------------
            # 前置作業，使用憑證文件，開啟 GoogleSheet
            # ------------------------------
            # 開啟 GoogleSheet 試算表(spreadsheet)
            shtInput  = self.gc.open_by_key('1a6Z1LSBV2u7P01X1vT0v5AyqqwcH-aj_dQ9pbfXV1X4') # 聰全資料表-選股記錄-成交量排行-20000
            shtOutput = self.gc.open_by_key('18kL5aIbNLudSUkgfu1pdGVYf-2citOQjjCST3lIInEA') # 選股記錄_2024
            # shtOutputDetails = self.gc.open_by_key('13ya5frtagmkCojhthWbG82L0NzTshX5PqwK25MNER3k') #技術指標明細

            # get GoogleSheet 中的第一個 工作表(worksheet)
            # wksInput0  = shtInput[0]
            wksInput1 = shtInput.worksheet_by_title('支持向量機-20000筆') #支持向量機
            wksInput2 = shtInput.worksheet_by_title('梯度提升樹-20000筆') #梯度提升樹
            wksOutput = shtOutput[0]
            # wksOutputDetails = shtOutputDetails[0]

            # 讀取「工作表(worksheet)」，存到 DataFrame
            dfInput1 = pd.DataFrame(wksInput1.get_all_records())
            dfInput2 = pd.DataFrame(wksInput2.get_all_records())
            print(F"\t dfInput1 size = {type(dfInput1)} {len(dfInput1)}")
            print(F"\t dfInput2 size = {type(dfInput2)} {len(dfInput2)}")


            # new class
            obj = StockUtils()

            # ------------------------------
            # get 「當天」內演算法選到的股票資料
            # ------------------------------
            # 原本想取一週內的全部重算，但是算到 71 筆的時候，就會被鎖，故還是改一日一日計算
            # get 今天
            # dToday = datetime.today()
            dToday = obj.adjustDate(datetime.today())
            sToday = dToday.strftime('%Y-%m-%d')
            dToday_date_only = dToday.date()

            # get 今天前7天的日期
            # dStartDate = dToday - timedelta(days=7)
            dStartDate = dToday - timedelta(days=0)
            print(F"\t dToday     = ({type(dToday)}) {dToday}")
            print(F"\t dStartDate = ({type(dStartDate)}) {dStartDate}")


            listStockCode = []
            setExist = set()
            # 宣告字典，記錄資料來源
            dictSourceSht = {}

            # # === 支持向量機 ===
            for row in dfInput1.itertuples(index=True, name='Pandas'):
                indexRow  = row.Index
                sDate = row[1]
                stockCode = row[2]
                stockName = row[3]
                # print(f"\t {self.sHr1}")
                # print(f"\t indexRow = {indexRow}, 資料日期 = {sDate}, 股票代號 = {stockCode}, 股票名稱 = {stockName}")

                # !!方便測試時，提早跳出迴圈
                # if indexRow == 1:
                if stockCode != 4916:
                    continue

                # 記錄資料來源
                key = f"{sDate}+{stockCode}"
                if key in dictSourceSht:
                    dictSourceSht[key].append('支持向量機')
                else:
                    dictSourceSht[key] = ['支持向量機']


                # 將 sDate 轉換為 datetime.date
                dDate = datetime.strptime(sDate, '%Y-%m-%d')
                dDate_date_only = dDate.date()

                # 檢查「資料日期+股票代號」如果不存在於 List 中，才將資料加入
                # if dStartDate <= dDate <= dToday and (sDate, stockCode) not in setExist:
                if dDate_date_only == dToday_date_only and (sDate, stockCode) not in setExist:
                    dict = {
                        "5日集中度": row[4],
                        "5日漲幅": row[5],
                        "20日漲幅": row[6],
                        "收盤價": row[7],
                        "漲幅(%)": row[8],
                        "外資連買(天)": row[9],
                        "外資連買張數": row[10],
                        "投信連買(天)": row[11],
                        "投信連買張數": row[12],
                        "自營商連買(天)": row[13],
                        "大戶近1週增減": row[14],
                        "散戶近1週增減％": row[15],
                        "近20日資餘增減": row[16],
                        "進出分點總家數差": row[17],
                        "上市上櫃": row[18],
                        "產業名稱": row[19],
                        "支持向量機實際運算數值": row[20],
                        "較5日均量縮放N%": row[21],
                        "27個技術指標看多": row[22],
                        "27個技術指標看空": row[23],
                        "最新股價": row[24],
                        "最新報酬率": row[25],
                        "觀察可以買的": row[26],
                    }
                    # 記錄到 List，用於後續產出
                    listStockCode.append([sDate, stockCode, stockName, dict])
                    # 如果「資料日期+股票代號」已記錄到 List ，不再增加到 List
                    setExist.add((sDate, stockCode))


            print(F"\t 1.listStockCode size = {type(listStockCode)} {len(listStockCode)}")

            # === 梯度提升樹 ===
            for row in dfInput2.itertuples(index=True, name='Pandas'):
                indexRow  = row.Index
                sDate = row[1]
                stockCode = row[2]
                stockName = row[3]
                # print(self.sHr1)
                # print(f"indexRow = {indexRow}, 資料日期 = {sDate}, 股票代號 = {stockCode}, 股票名稱 = {stockName}")

                # 記錄資料來源
                key = f"{sDate}+{stockCode}"
                if key in dictSourceSht:
                    dictSourceSht[key].append('梯度提升樹')
                else:
                    dictSourceSht[key] = ['梯度提升樹']


                # 將 sDate 轉換為 datetime.date
                dDate = datetime.strptime(sDate, '%Y-%m-%d')
                dDate_date_only = dDate.date()

                # 檢查「資料日期+股票代號」如果不存在於 List 中，才將資料加入
                # if dStartDate <= dDate <= dToday and (sDate, stockCode) not in setExist:
                if dDate_date_only == dToday_date_only and (sDate, stockCode) not in setExist:
                    dict = {
                        "5日集中度": row[4],
                        "5日漲幅": row[5],
                        "20日漲幅": row[6],
                        "收盤價": row[7],
                        "漲幅(%)": row[8],
                        "外資連買(天)": row[9],
                        "外資連買張數": row[10],
                        "投信連買(天)": row[11],
                        "投信連買張數": row[12],
                        "自營商連買(天)": row[13],
                        "大戶近1週增減": row[14],
                        "散戶近1週增減％": row[15],
                        "近20日資餘增減": row[16],
                        "進出分點總家數差": row[17],
                        "上市上櫃": row[18],
                        "產業名稱": row[19],
                        "較5日均量縮放N%": row[20],
                        "27個技術指標看多": row[21],
                        "27個技術指標看空": row[22],
                        "多頭指標連3天增加": row[23],
                        "最新股價": row[24],
                        "最新報酬率": row[25],
                        "觀察可以買的": row[26],
                    }
                    # 記錄到 List，用於後續產出
                    listStockCode.append([sDate, stockCode, stockName, dict])
                    # 如果「資料日期+股票代號」已記錄到 List ，不再增加到 List
                    setExist.add((sDate, stockCode))
            print(F"\t 2.listStockCode size = {type(listStockCode)} {len(listStockCode)}")


            # ------------------------------
            # 組出要寫到 GoogleSheet 的內容
            # ------------------------------
            listOutput = []
            for i in range(len(listStockCode)):
                listCurrent = listStockCode[i]

                # !!方便測試時，提早跳出迴圈
                # if i == 1:
                #     break

                stockDate = listCurrent[0] # 資料日期
                stockCode = listCurrent[1] # 股票代號
                stockName = listCurrent[2] # 股票名稱
                dict      = listCurrent[3] # 詳細數據
                print(f"\t {self.sHr1}")
                print(f"\t i = {i}, 資料日期 = {stockDate}, 股票代號 = {stockCode}, 股票名稱 = {stockName}")

                # string 轉成 datetime.datetime
                dStockDate = datetime.strptime(stockDate, '%Y-%m-%d')
                # 「抓取資料的開始日期」預設為 今日-60
                dStockDate = dStockDate - timedelta(days=60)

                # 取得各技術指標
                sEndDate = stockDate
                sStartDate = dStockDate.strftime('%Y-%m-%d')
                listStockInfo = obj.getStockAnalyze(str(stockCode),sEndDate,sStartDate)
                print(F"\t listStockInfo size = {type(listStockInfo)} {len(listStockInfo)}")

                for i in range(len(listStockInfo)):
                    listCurrentDay = listStockInfo [i]

                    sDate = listCurrentDay[0]  # 資料日期

                    # 只取當天的資料，故日期非當天即跳出迴圈
                    if (sDate != stockDate):
                        break

                    stockCode     = listCurrentDay[1] # 股票代號
                    stockName     = listCurrentDay[2] # 股票名稱
                    fClosePrice   = listCurrentDay[3] # 收盤價
                    iBullishCount = listCurrentDay[4] # 多頭數量
                    iBearishCount = listCurrentDay[5] # 空頭數量
                    # dictSignals   = listCurrentDay[6] # 詳細各計術指標值、歷史價格數據

                    # 取得「多頭數量」連續增加天數
                    iUpDays   = obj.countIncreaseDays("up",listStockInfo)
                    # 取得「空頭數量」連續增加天數
                    iDownDays = obj.countIncreaseDays("down",listStockInfo)
                    # print(F"\t iUpDays   = ({type(iUpDays)}) {iUpDays}")
                    print(F"\t iDownDays = ({type(iDownDays)}) {iDownDays}")


                    # 取得「多頭數量」連續 >=index 的天數
                    iUpDays20 = obj.countIndexDays('up', listStockInfo, 20)
                    iUpDays15 = obj.countIndexDays('up', listStockInfo, 15)
                    # print(F"\t iUpDays20 = ({type(iUpDays20)}) {iUpDays20}")
                    # print(F"\t iUpDays15 = ({type(iUpDays15)}) {iUpDays15}")

                    # 取得「空頭數量」連續 >=index 的天數
                    iDownDays20 = obj.countIndexDays('down', listStockInfo, 20)
                    iDownDays15 = obj.countIndexDays('down', listStockInfo, 15)
                    # print(F"\t iDownDays20 = ({type(iDownDays20)}) {iDownDays20}")
                    # print(F"\t iDownDays15 = ({type(iDownDays15)}) {iDownDays15}")

                    # 將數據保存到 Excel 文件
                    obj.saveToExcel(listStockInfo,"C:/Users/user/Desktop/股票技術指標輸出")

                    # 將數據保存到 Google 文件
                    obj.saveToGoogle(listStockInfo)

                    # ------------------------------
                    # get 買進Flag
                    # ------------------------------
                    # 5日漲幅 < 5
                    # 20日漲幅 < 10
                    # 外資連買(天) > 1 OR 投信連買(天) > 1
                    # 進出分點總家數差 < 0
                    # 「多頭數量」連續增加天數 >= 3 OR 取得「多頭數量」連續大於20天數 > 3
                    growth05 = dict.get("5日漲幅", 0)
                    growth20 = dict.get("20日漲幅", 0)
                    foreign_investors_buying = dict.get("外資連買(天)", 0)
                    investment_trust_buying  = dict.get("投信連買(天)", 0)
                    households = dict.get("進出分點總家數差", 0)
                    # print(F"\t growth05 = ({type(growth05)}) {growth05}")
                    # print(F"\t growth20 = ({type(growth20)}) {growth20}")
                    # print(F"\t foreign_investors_buying = ({type(foreign_investors_buying)}) {foreign_investors_buying}")
                    # print(F"\t investment_trust_buying = ({type(investment_trust_buying)}) {investment_trust_buying}")
                    # print(F"\t households = ({type(households)}) {households}")


                    isBuy = False
                    if growth05 < 5 and growth20 < 10:
                        if (foreign_investors_buying > 1 or investment_trust_buying > 1):
                            if households < 0:
                                if iUpDays >= 3 or iUpDays20 > 3:
                                    isBuy = True


                    key = f"{sDate}+{stockCode}"
                    listOutput.append({
                        "更新日期": sDate,
                        "股票代號": stockCode,
                        "股票名稱": stockName,
                        "買進Flag": isBuy,
                        "最新股價": "", #"=GOOGLEFINANCE(\"TPE:\"&B2}, \"price\")",
                        "最新報酬率": "", #"=(E2-G2)/G2*100",
                        "收盤價": fClosePrice,
                        "漲幅(%)": dict.get("漲幅(%)", 0),
                        "技術指標-多頭數量" : iBullishCount,
                        "技術指標-空頭數量" : iBearishCount,
                        "「多頭數量」連續增加天數" : iUpDays,
                        "「多頭數量」連續>20天數"  : iUpDays20,
                        "「多頭數量」連續>15天數"  : iUpDays15,
                        "「空頭數量」連續增加天數" : iDownDays,
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
                        "資料來源" : dictSourceSht.get(key, 0),
                    })


            # ------------------------------
            # 將資料寫入 GoogleSheet 的內容
            # ------------------------------
            # Set 表頭
            columns = ["更新日期","股票代號","股票名稱"
                        ,"買進Flag","最新股價","最新報酬率","收盤價","漲幅(%)"
                        ,"技術指標-多頭數量","技術指標-空頭數量"
                        ,"「多頭數量」連續增加天數","「多頭數量」連續>20天數","「多頭數量」連續>15天數"
                        ,"「空頭數量」連續增加天數","「空頭數量」連續>20天數","「空頭數量」連續>15天數"
                        ,"5日漲幅","20日漲幅"
                        ,"外資連買(天)","外資連買張數","投信連買(天)","投信連買張數","自營商連買(天)"
                        ,"大戶近1週增減","散戶近1週增減％","進出分點總家數差","資料來源"]
            print(F"\t columns size {type(columns)} {len(columns)}")
            print(F"\t listOutput size {type(listOutput)} {len(listOutput)}")

            dfOutput = pd.DataFrame(listOutput, columns=columns)

            # 將資料依「技術指標-多頭數量」排序
            dfOutput = dfOutput.sort_values(by='技術指標-多頭數量', ascending=False)

            # get 目前工作表資料，並轉為 DataFrame
            existingData = wksOutput.get_all_records()
            dfExisting = pd.DataFrame(existingData)

            # 將新的資料附加到現有資料後面
            dfCombined = pd.concat([dfExisting, dfOutput], ignore_index=True)

            # 將更新後的 DataFrame 寫回工作表
            wksOutput.set_dataframe(dfCombined, (0, 0))

            # 設置公式
            # formulas = [
            #     ('E1', '=GOOGLEFINANCE("TPE:"&B2, "price")'),  # 最新股價
            #     ('F1', '=(E2-G2)/G2*100')  # 最新報酬率
            # ]

            # for cell, formula in formulas:
            #     wksOutput.update_value(cell, formula)


        except Exception as e:
            print(f"An error occurred: {e}")

# ==================================================
# new class
obj = StockSaveToGoogle()

obj.upBuyingStock()