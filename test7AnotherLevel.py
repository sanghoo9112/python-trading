import win32com.client

def PrintCalList(instMarketEye, code):

    # 4:현재가(long or float)
    # 62:외국인순매매(long)
    # 67:PER(float)
    # 75:부채비율(float)
    # 78:매출액증가율(float)
    # 80:순이익증가율(float)
    # 88:당기순이익(ulonglog) - 단위:원
    # 97:분기매출액증가율(float)
    # 110:분기부채비율(float)
    instMarketEye.SetInputValue(0, (4, 62, 67, 75, 78, 80, 88, 97, 110))
    instMarketEye.SetInputValue(1, electCodeList)

    # BlockRequest
    instMarketEye.BlockRequest()

    # GetHeaderValue
    numStock = instMarketEye.GetHeaderValue(2)

    for i in range(numStock):
        print("----------------------------------------")
        print("PER: ", instMarketEye.GetDataValue(0, i))
        print("EPS: ", instMarketEye.GetDataValue(1, i))
        print("최근분기년월: ", instMarketEye.GetDataValue(2, i))
        print("----------------------------------------")

if __name__ == "__main__":
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

    # 13 전기전자
    # 145 정보기술

    print(instCpCodeMgr.GetIndustryName(13))
    print("정보")
    electCodeList = instCpCodeMgr.GetGroupCodeList(13)

    for code in range(electCodeList):
        print("-----------------------------------------")
        print(code, instCpCodeMgr.CodeToName(code))
        PrintCalList(instCpCodeMgr,code)
        print("-----------------------------------------")






