import win32com.client

instCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
print(instCpCybos.IsConnect)

instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
stockNum = instCpStockCode.GetCount()


naverCode = instCpStockCode.NameToCode('NAVER')
naverIndex = instCpStockCode.CodeToIndex(naverCode)
print(naverCode)
print(naverIndex)


#for i in range(0,10):
    #print(instCpStockCode.GetData(1, i))