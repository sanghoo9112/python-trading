import win32com.client

def PrintCalList(instCpCodeMgr,codeList):

    # 4:현재가(long or float)
    # 62:외국인순매매(long)
    # 67:PER(float)
    # 75:부채비율(float)
    # 78:매출액증가율(float)
    # 80:순이익증가율(float)
    # 88:당기순이익(ulonglog) - 단위:원
    # 97:분기매출액증가율(float)
    # 110:분기부채비율(float)

    print("전기전자")
    for code1 in electCodeList:



        print(code1, instCpCodeMgr.CodeToName(code1))

    print("-------------------------")
    print("정보기술")

    for code2 in itCodeList:
        print(code2, instCpCodeMgr.CodeToName(code2))


if __name__ == "__main__":
    instCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    instMarketEye = win32com.client.Dispatch("CpSysDib.MarketEye")

    # 13 전기전자
    # 145 정보기술

    # 전기전자 코드리스트 남아오고
    electCodeList = instCpCodeMgr.GetGroupCodeList(13)
    # 정보기술 코드리스트를 담아온다
    itCodeList = instCpCodeMgr.GetGroupCodeList(145)

    PrintCalList()





