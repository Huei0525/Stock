from StockUtils import StockUtils

# new class
obj = StockUtils()

# 日期區間測試
# istStockInfo = obj.getStockAnalyzeTEST("3703","2024-07-10","") #欣陸
# istStockInfo = obj.getStockAnalyzeTEST("3703","2024-07-11","") #欣陸


# 取得各技術指標
listStockInfo = obj.getStockAnalyze("3703","2024-07-13","2024-05-24") #欣陸
# listStockInfo = obj.getStockAnalyze("3532","2024-06-21","2024-04-23") #台勝科
# listStockInfo = obj.getStockAnalyze("6197","2024-06-21","2024-04-23") #佳必琪
# listStockInfo = obj.stockAnalyze("2421","2024-06-21","2024-04-23") #建準
# print(F"\t listStockInfo size = {type(listStockInfo)} {len(listStockInfo)}")

# listStockInfo = [
#     ['2024-06-26', '5284', 'jpp-KY', 128.5, 16, 9],
#     ['2024-06-25', '5284', 'jpp-KY', 127.5, 15, 8],
#     ['2024-06-24', '5284', 'jpp-KY', 126.0, 15, 12],
#     ['2024-06-21', '5284', 'jpp-KY', 126.5, 13, 10],
#     ['2024-06-20', '5284', 'jpp-KY', 125.5, 11, 12],
    # ['2024-06-19', '5284', 'jpp-KY', 126.0, 12, 11],
    # ['2024-06-18', '5284', 'jpp-KY', 127.5, 16, 7],
    # ['2024-06-17', '5284', 'jpp-KY', 128.0, 18, 5],
    # ['2024-06-14', '5284', 'jpp-KY', 123.0, 16, 7],
    # ['2024-06-13', '5284', 'jpp-KY', 117.0, 4, 19],
    # ['2024-06-12', '5284', 'jpp-KY', 117.5, 4, 19],
    # ['2024-06-11', '5284', 'jpp-KY', 117.0, 2, 21],
    # ['2024-06-07', '5284', 'jpp-KY', 121.5, 7, 16],
    # ['2024-06-06', '5284', 'jpp-KY', 118.0, 7, 16],
    # ['2024-06-05', '5284', 'jpp-KY', 125.5, 16, 7],
    # ['2024-06-04', '5284', 'jpp-KY', 124.0, 6, 17],
    # ['2024-06-03', '5284', 'jpp-KY', 124.0, 6, 17],
    # ['2024-05-31', '5284', 'jpp-KY', 123.5, 4, 19],
    # ['2024-05-30', '5284', 'jpp-KY', 124.0, 6, 17],
    # ['2024-05-29', '5284', 'jpp-KY', 127.0, 16, 7],
    # ['2024-05-28', '5284', 'jpp-KY', 126.0, 11, 12],
    # ['2024-05-27', '5284', 'jpp-KY', 123.0, 6, 17],
    # ['2024-05-24', '5284', 'jpp-KY', 122.5, 5, 18],
    # ['2024-05-23', '5284', 'jpp-KY', 121.5, 8, 15],
    # ['2024-05-22', '5284', 'jpp-KY', 123.5, 5, 18],
    # ['2024-05-21', '5284', 'jpp-KY', 122.5, 2, 21],
    # ['2024-05-20', '5284', 'jpp-KY', 123.5, 2, 21],
    # ['2024-05-17', '5284', 'jpp-KY', 124.0, 2, 21],
    # ['2024-05-16', '5284', 'jpp-KY', 124.5, 3, 20],
    # ['2024-05-15', '5284', 'jpp-KY', 126.0, 4, 19],
    # ['2024-05-14', '5284', 'jpp-KY', 129.5, 7, 16]
# ]

# print("==================================================")
# 取得「多頭數量」連續增加天數
# iUpDays   = obj.countIncreaseDays("up",listStockInfo)
# # 取得「空頭數量」連續增加天數
# iDownDays = obj.countIncreaseDays("down",listStockInfo)
# print(F"iUpDays   = ({type(iUpDays)}) {iUpDays}")
# print(F"iDownDays = ({type(iDownDays)}) {iDownDays}")

# dictDays = obj.countIncreaseDays(listStockInfo)
# print(F"dictDays = ({len(dictDays)}) {dictDays}")

# print("==================================================")
# # 取得「多頭數量」連續 >=index 的天數
# iUpDays20 = obj.countIndexDays('up', listStockInfo, 20)
# iUpDays15 = obj.countIndexDays('up', listStockInfo, 15)
# print(F"\t iUpDays20 = ({type(iUpDays20)}) {iUpDays20}")
# print(F"\t iUpDays15 = ({type(iUpDays15)}) {iUpDays15}")

# # 取得「空頭數量」連續 >=index 的天數
# iDownDays20 = obj.countIndexDays('down', listStockInfo, 20)
# iDownDays15 = obj.countIndexDays('down', listStockInfo, 15)
# print(F"\t iDownDays20 = ({type(iDownDays20)}) {iDownDays20}")
# print(F"\t iDownDays15 = ({type(iDownDays15)}) {iDownDays15}")


# print("==================================================")
# 針對傳入的list，每一天都取得「多頭/空頭數量」連續增加天數，回傳 dict
# dictDays = obj.countIncreaseDaysHistory(listStockInfo)


# print("==================================================")
# 將數據保存到 Excel 文件中
# obj.saveToExcel("C:/Users/user/Desktop/股票技術指標輸出",listStockInfo,dictDays)


# print("==================================================")
# 將數據保存到 Google 文件
# obj.saveToGoogle(listStockInfo)


# print("==================================================")
# 爬Yahoo網頁,取得當天股票資訊
# dict = obj.getStockData4Yahoo("3532")
# print(F"dict = ({type(dict)}) {dict}")


# print("==================================================")
# 使用TWStock,取得股票資訊

# listStockInfo = obj.getStockAnalyze("5284","2024-06-26","2024-05-14") #jpp-KY
# obj.getStockData4TWStock("5284")

