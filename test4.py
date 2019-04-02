import win32com.client

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instStockChart.SetInputValue(0, "A003540")
# 개수를 요청할때는 2를 입력하고 기간을 입력할때는 1을 입력한다.
instStockChart.SetInputValue(1, ord('2'))
# 요청 개수랑는 타입 4 ->10 이 실제로 요청할 데이터의 개수를 의미한다.
instStockChart.SetInputValue(4, 10)
# 요청할 데이터의 종류 5 -> 종가를 의미한다.
instStockChart.SetInputValue(5, 5)
# 차트의 종류로서 데이터를 가져오기 위해 D 를 입력했다.
instStockChart.SetInputValue(6, ord('D'))
#
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()

# 데이터 원하는걸 입력 후 데이터 처리를 요청하면 된다.
numData = instStockChart.GetHeaderValue(3)
for i in range(numData):
    print(instStockChart.GetDataValue(0, i))

