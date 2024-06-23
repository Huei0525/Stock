from StockUtils import StockUtils

# new class
obj = StockUtils()

# 取得各技術指標
listStockInfo = obj.getStockAnalyze("1104","2024-06-18","2024-04-18") #環泥
# listStockInfo = obj.getStockAnalyze("3532","2024-06-21","2024-04-23") #台勝科
# listStockInfo = obj.getStockAnalyze("6197","2024-06-21","2024-04-23") #佳必琪
# listStockInfo = obj.stockAnalyze("2421","2024-06-21","2024-04-23") #建準
# print(F"\t listStockInfo size = {type(listStockInfo)} {len(listStockInfo)}")


print("==================================================")
# 取得「多頭數量」連續增加天數
iUpDays   = obj.countIncreaseDays("up",listStockInfo)
# 取得「空頭數量」連續增加天數
iDownDays = obj.countIncreaseDays("down",listStockInfo)
print(F"iUpDays   = ({type(iUpDays)}) {iUpDays}")
print(F"iDownDays = ({type(iDownDays)}) {iDownDays}")


print("==================================================")
# 取得「多頭數量」連續 >=index 的天數
iUpDays20 = obj.countIndexDays('up', listStockInfo, 20)
iUpDays15 = obj.countIndexDays('up', listStockInfo, 15)
print(F"\t iUpDays20 = ({type(iUpDays20)}) {iUpDays20}")
print(F"\t iUpDays15 = ({type(iUpDays15)}) {iUpDays15}")

# 取得「空頭數量」連續 >=index 的天數
iDownDays20 = obj.countIndexDays('down', listStockInfo, 20)
iDownDays15 = obj.countIndexDays('down', listStockInfo, 15)
print(F"\t iDownDays20 = ({type(iDownDays20)}) {iDownDays20}")
print(F"\t iDownDays15 = ({type(iDownDays15)}) {iDownDays15}")


print("==================================================")
# 將數據保存到 Excel 文件
obj.saveToExcel(listStockInfo,"C:/Users/user/Desktop/股票技術指標輸出")


print("==================================================")
# 將數據保存到 Google 文件
# obj.saveToGoogle(listStockInfo)



print("==================================================")
# 爬Yahoo網頁,取得當天股票資訊
# dict = obj.getStockData4Yahoo("3532")
# print(F"dict = ({type(dict)}) {dict}")