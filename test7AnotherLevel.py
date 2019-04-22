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
    #instMarketEye.SetInputValue(0, (4, 67, 70, 111))
    instMarketEye.SetInputValue(1, code)

    # BlockRequest
    instMarketEye.BlockRequest()
    print("현재가(long or float) : ", instMarketEye.GetDataValue(0, 0))
    print("외국인순매매(long): ", instMarketEye.GetDataValue(1, 0))
    print("PER(float): ", instMarketEye.GetDataValue(2, 0))
    print("부채비율(float): ", instMarketEye.GetDataValue(3, 0))
    print("매출액증가율(float): ", instMarketEye.GetDataValue(4, 0))
    print("당기순이익(ulonglog) - 단위:원: ", instMarketEye.GetDataValue(5, 0))
    print("분기매출액증가율(float): ", instMarketEye.GetDataValue(6, 0))
    print("분기부채비율(float): ", instMarketEye.GetDataValue(7, 0))


if __name__ == "__main__":
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

    # 13 전기전자
    # 145 정보기술

    print(instCpCodeMgr.GetIndustryName(145))
    print("정보")
    electCodeList = instCpCodeMgr.GetGroupCodeList(145)

    for i, code in enumerate(electCodeList):
        print("-----------------------------------------")
        print(code, instCpCodeMgr.CodeToName(code))
        PrintCalList(instMarketEye, code)






