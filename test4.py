import win32com.client

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instStockChart.SetInputValue(0, "A003540")
# 개수를 요청할때는 2를 입력하고 기간을 입력할때는 1을 입력한다.
instStockChart.SetInputValue(1, ord('2'))
# 요청 개수랑는 타입 4 ->10 이 실제로 요청할 데이터의 개수 (타입 = 데이터의 갯수 ) 를 의미한다.
instStockChart.SetInputValue(4, 10)
# 요청할 데이터의 종류 5 -> 종가를 의미한다. 0 : 날짜 / 1: 시간 / 2:시가 / 4: 고가 / 5: 종가 / 8 : 거래량
instStockChart.SetInputValue(5, (0, 2, 3, 4, 5, 8))
# 차트의 종류로서 데이터를 가져오기 위해 D 를 입력했다.
instStockChart.SetInputValue(6, ord('D'))
# 수정증가 무수정증가?
instStockChart.SetInputValue(9, ord('1'))

instStockChart.BlockRequest()

# 데이터 원하는걸 입력 후 데이터 처리를 요청하면 된다.
numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)


for i in range(numData):
    for j in range(numField):
        print(instStockChart.GetDataValue(j, i), end=" ")
    print("")

