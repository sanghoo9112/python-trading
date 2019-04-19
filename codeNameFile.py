import win32com.client

instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = instCpCodeMgr.GetStockListByMarket(1)
print(codeList)


kospi = {}
for code in codeList:
    name = instCpCodeMgr.CodeToName(code)
    kospi[code] = name


f = open('C:\\Users\\WAUG-PC-73\\Desktop\\kospi.csv', 'w')
for key, value in kospi.items():
    f.write("%s,%s\n" % (key, value))
f.close()
