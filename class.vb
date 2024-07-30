Public Function ReadSSPersonInfo() As String
    DownloadRegistration
    ReadSSPersonInfo = ReadSSCard(True)
    Exit Function
End Function
Public Function ReadSSMagCard() As String
    DownloadRegistration
    ReadSSMagCard = ReadSSCard(False)
    Exit Function
End Function
Public Function ReadSSCard(detailFlag As Boolean) As String
    DownloadRegistration
    Dim iHandle As Integer
    Dim Rtn As Integer
    Dim iType As Integer
    Dim iDNo As String
    Dim iReturn As Integer
    
    For iPort = 100 To 200
        iHandle = dc_init(iPort, 0)
        If iHandle > 0 Then
            Exit For
        End If
    Next iPort
    If iHandle > 0 Then
        Rtn = dc_beep(iHandle, 10)
        If (Rtn <> 0) Then
            ReadSSCard = "-1^" & "设备蜂鸣失败"
            'Exit Function
        End If

        Dim CardIdentifier As String
        Dim CardType_ As String
        Dim CardVersion As String
        Dim IssuersID As String
        Dim IssuingDate As String
        Dim EffectiveDate As String
        Dim InsuCardNo As String
        Dim CardID As String
        Dim name As String
        Dim name_ As String
        Dim sex As String
        Dim nation As String
        Dim address As String
        Dim Birthday As String
        
        CardIdentifier = String$(255, Chr$(0))
        CardType_ = String$(255, Chr$(0))
        CardVersion = String$(255, Chr$(0))
        IssuersID = String$(255, Chr$(0))
        IssuingDate = String$(255, Chr$(0))
        EffectiveDate = String$(255, Chr$(0))
        InsuCardNo = String$(255, Chr$(0))
        CardID = String$(255, Chr$(0))
        name = String$(255, Chr$(0))
        name_ = String$(255, Chr$(0))
        sex = String$(255, Chr$(0))
        nation = String$(255, Chr$(0))
        Birthday = String$(255, Chr$(0))
        address = String$(255, Chr$(0))
        age = String$(255, Chr$(0))
        
        iType = 1
        iReturn = dc_GetSocialSecurityCardBaseInfo(iHandle, iType, CardIdentifier, CardType_, CardVersion, IssuersID, IssuingDate, EffectiveDate, InsuCardNo, CardID, name, name_, sex, nation, address, Birthday)
        If (iReturn <> 0) Then
            iReturn = dc_exit(iPort)
            ReadSSCard = "-1^社保卡读取信息失败！错误代码：" & iReturn
            dc_exit (iPort)
            Exit Function
        End If
        
        dc_exit (iPort)
        
        '(CardIdentifier, CardType,CardVersion,IssuersID,IssuingDate,EffectiveDate,InsuCardNo)
        CardType_ = Replace(CardType_, Chr(0), "")
        CardVersion = Replace(CardVersion, Chr(0), "")
        IssuersID = Replace(IssuersID, Chr(0), "")
        IssuingDate = Replace(IssuingDate, Chr(0), "")
        EffectiveDate = Replace(EffectiveDate, Chr(0), "")
        InsuCardNo = Replace(InsuCardNo, Chr(0), "")
        CardIdentifier = Replace(CardIdentifier, Chr(0), "")
        
        
        CardID = Replace(CardID, Chr(0), "")
        name = Replace(name, Chr(0), "")
        name_ = Replace(name_, " ", "")
        sex = Replace(sex, Chr(0), "")
        nation = Replace(nation, Chr(0), "")
        Birthday = Replace(Birthday, Chr(0), "")
        Birthday = Mid(Birthday, 1, 4) & "-" & Mid(Birthday, 5, 2) & "-" & Mid(Birthday, 7, 2)
        address = Replace(address, Chr(0), "")
        age = "" 'GetAge(IDNo)
        
        myXMLData = myXMLData & GetXMLNodeData(gPeopleName, name)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleSex, sex)
        myXMLData = myXMLData & GetXMLNodeData("CredNo", CardID)
        myXMLData = myXMLData & GetXMLNodeData("NationDescLookUpRowID", nation)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleNation, nation)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleBirthday, Birthday)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleAddress, address)
        myXMLData = myXMLData & GetXMLNodeData("CardNo", InsuCardNo)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleAge, "")
        myXMLData = myXMLData & GetXMLNodeData(gInsuCardNo, "")
        myXMLData = myXMLData & GetXMLNodeData("CardIdentifier", CardIdentifier)
        myXMLData = myXMLData & GetXMLNodeData("CardType", CardType_)
        myXMLData = myXMLData & GetXMLNodeData("CardVersion", CardVersion)
        myXMLData = myXMLData & GetXMLNodeData("IssuersID", IssuersID)
        myXMLData = myXMLData & GetXMLNodeData("IssuingDate", IssuingDate)
        myXMLData = myXMLData & GetXMLNodeData("EffectiveDate", EffectiveDate)
        myXMLData = "<" & gRoot & ">" & myXMLData & "</" & gRoot & ">"
        
        InsuCardInfo = CheckNullChar(InsuCardNo)
        
        If (detailFlag) Then
            ReadSSCard = "0^" & myXMLData
            Exit Function
        Else
            ReadSSCard = "0^" & CardID & "^^" & myXMLData
            Exit Function
        End If
    End If
    Exit Function
End Function
Public Function ReadPersonInfo() As String
    DownloadRegistration
    ReadPersonInfo = ReadIDCard(True)
    Exit Function
End Function
Public Function ReadMagCard() As String
    DownloadRegistration
    ReadMagCard = ReadIDCard(False)
    Exit Function
End Function
Public Function ReadIDCard(detailFlag As Boolean) As String
    DownloadRegistration
    Dim lRet As Integer
    Dim TRet As Integer
    Dim iType As Integer
    Dim iReturn As Integer
    Dim Rtn As Integer
    
    Dim name As String
    Dim sex As String
    Dim nation As String
    Dim birth As String
    Dim address As String
    Dim iDNo As String
    Dim age As String
    Dim department As String
    Dim expire_start_day As String
    Dim expire_end_day As String
    Dim reserved As String
    Dim myXMLData As String
    
    Dim name_(1024) As Byte
    Dim sex_(1024) As Byte
    Dim nation_(1024) As Byte
    Dim birth_(1024) As Byte
    Dim address_(1024) As Byte
    Dim iDNo_(1024) As Byte
    Dim age_(1024) As Byte
    Dim department_(1024) As Byte
    Dim expire_start_day_(1024) As Byte
    Dim expire_end_day_(1024) As Byte
    Dim reserved_(1024) As Byte
    
    Dim MsgLen As Long
    Dim PhotoLen As Long
    Dim FingerLen As Long
    Dim ExtraLen As Long
    Dim Base64Len As Long
    Dim pMsg(1024) As Byte
    Dim photo(1024) As Byte
    Dim finger(1024) As Byte
    Dim extra(70) As Byte
    Dim base64(65536) As Byte
    
    
    name = String$(255, Chr$(0))
    sex = String$(255, Chr$(0))
    nation = String$(255, Chr$(0))
    birth = String$(255, Chr$(0))
    address = String$(255, Chr$(0))
    iDNo = String$(255, Chr$(0))
    age = String$(255, Chr$(0))
    department = String$(255, Chr$(0))
    expire_start_day = String$(255, Chr$(0))
    expire_end_day = String$(255, Chr$(0))
    reserved = String$(255, Chr$(0))
    
    For iPort = 100 To 200
        iHandle = dc_init(iPort, 0)
        If iHandle > 0 Then
            Exit For
        End If
    Next iPort
    If iHandle > 0 Then
        Rtn = dc_beep(iHandle, 20)
        If Rtn <> 0 Then
            iReturn = dc_exit(iPort)
            ReadIDCard = "-1^" & "设备蜂鸣失败"
            Exit Function
        End If
        'int MsgLen = 0, PhotoLen = 0, FingerLen = 0, ExtraLen = 0, Base64Len = 65536;
        MsgLen = 0
        PhotoLen = 0
        FingerLen = 0
        ExtraLen = 0
        lRet = dc_SamAReadCardInfo(iHandle, 1, MsgLen, pMsg(0), PhotoLen, photo(0), FingerLen, finger(0), ExtraLen, extra(0))
        If (lRet <> 0) Then
            iReturn = dc_exit(iPort)
            ReadIDCard = "-1^" & "读取身份证信息失败,请确认身份证是否放置,身份证是否有效"
            Exit Function
        End If
        
        '返回1表示社保卡；返回2表示居民健康卡，返回3表示M1卡， 返回4表示二代证，返回5表示银行卡，返回6表示无卡
        iType = dc_GetIdCardType(iHandle, MsgLen, pMsg)
        
        If (iType = 0) Then
            Rtn = dc_ParseTextInfo(iHandle, 0, MsgLen, pMsg(0), name_(0), sex_(0), nation_(0), birth_(0), address_(0), iDNo_(0), department_(0), expire_start_day_(0), expire_end_day_(0), reserved_(0))
            If (Rtn <> 0) Then
                ReadIDCard = "-1^读取身份证信息失败"
                Exit Function
            End If
        End If
        
        '老外国人居留证
        If (iType = 1) Then
            'Rtn = dc_ParseTextInfoForForeigner(iHandle, 0, MsgLen, pMsg(0), name_(0), sex_(0), nation_(0), birth_(0), address_(0), iDNo_(0), department_(0), expire_start_day_(0), expire_end_day_(0), reserved_(0))
            'Exit Function
        End If
        
        '新外国人居留证
        If (iType = 3) Then
            
            'Rtn = dc_ParseTextInfoForNewForeigner(iHandle, 0, MsgLen, pMsg(0), name_(0), sex_(0), nation_(0), birth_(0), address_(0), iDNo_(0), department_(0), expire_start_day_(0), expire_end_day_(0), reserved_(0))
            'Exit Function
        End If
        
        dc_exit (iPort)
        
        name = ByteArrayToString(name_)
        sex = ByteArrayToString(sex_)
        nation = ByteArrayToString(nation_)
        birth = ByteArrayToString(birth_)
        address = ByteArrayToString(address_)
        iDNo = ByteArrayToString(iDNo_)
        department = ByteArrayToString(department_)
        expire_start_day = ByteArrayToString(expire_start_day_)
        expire_end_day = ByteArrayToString(expire_end_day_)
        reserved = ByteArrayToString(reserved_)
        
        sex = getSexCode(sex)
        nation = Replace(nation, "0", "")
        'nation = getNationCode(nation)
        birth = Mid(birth, 1, 4) & "-" & Mid(birth, 5, 2) & "-" & Mid(birth, 7, 2)
        age = ""
        'birth = Trim(TrimASCII(birth_))
        'address = Trim(TrimASCII(address_))
        'iDNo = Trim(TrimASCII(iDNo_))
        
        'photo
        
        
        myXMLData = myXMLData & GetXMLNodeData(gPeopleName, name)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleSex, sex)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleNation, nation)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleBirthday, birth)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleAddress, address)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleIDCode, iDNo)
        myXMLData = myXMLData & GetXMLNodeData(gPeopleAge, age)
        'myXMLData = myXMLData & GetXMLNodeData(gInsuCardNo, InsuCardNo)
        'myXMLData = myXMLData & GetXMLNodeData(gCName, CName)
        myXMLData = "<" & gRoot & ">" & myXMLData & "</" & gRoot & ">"
        
        
        
        If (detailFlag) Then
            ReadIDCard = "0^" & myXMLData
            Exit Function
        Else
            ReadIDCard = "0^" & iDNo & "^^" & myXMLData
            Exit Function
        End If
    Else
        Exit Function
    End If
End Function
'读身份证
Public Function ReadIDCardNo(CardType As String) As String

    Dim Rtn As Integer
    Dim m1str As String
    Dim Person As String
   Call InitInfo
   'Device = "SS"
   If (Device = "DK") Then
        m1str = ReadIDCardNoDK(CardType)
    Else
        m1str = ReadIDCardNoSS(CardType)
    End If
    ReadIDCardNo = m1str

End Function
Public Function ReadIDCardNoDK(CardType As String) As String
    DownloadRegistration
    Dim lRet As Integer
    Dim Rtn As Integer
    Dim iType As Integer
    Dim pchOutInfo As String * 20480
    Dim iDNo As String
    
    
    Dim pErrMsg(64) As Byte
    Rtn = DeviceBeep(10, pErrMsg(0))
    If Rtn <> 0 Then
        ReadIDCardNoDK = "-1^" & "设备蜂鸣失败"
        Exit Function
    End If
    iType = iCheckCardType()
    If (iType = 2) Or (iType = 5) Or (iType = 6) Then
        ReadIDCardNoDK = "-100^卡类型错误"
        Exit Function
    End If
    lRet = iReadIdentityCard(1, pchOutInfo, 2)
    If (lRet < 0) Then
          ReadIDCardNoDK = lRet & "^" & pchOutInfo
          Exit Function
    End If
        
    pchOutInfo = StrConv(pchOutInfo, vbNarrow)
    PatArr = Split(pchOutInfo, "|")
    iDNo = PatArr(5)
    iDNo = Trim(TrimASCII(iDNo))
    ReadIDCardNoDK = "0" & "^" & iDNo & "^" & "" & "^"
    Exit Function
End Function
Public Function ReadIDCardNoSS(CardType As String) As String
    DownloadRegistration
    Dim handle As Long
    Dim iReturn As Long
    Dim IDCard As String, name As String, sex As String, Folk As String, Brith As String, address As String
    Dim ReadNum As Integer, ReadStatus As Integer
    iReturn = 0
    ReadStatus = 0
    ReadNum = 5
    IDCard = String$(255, Chr$(0))
    handle = ss_reader_open()
    If handle <= 0 Then
        ReadIDCardNoSS = "-1^设备打开失败，错误代码：" & CStr(handle)
        Exit Function
    End If
'    Do While (ReadNum > 0 And ReadStatus = 0)
        iReturn = ss_id_ResetID2Card(handle)
        If iReturn <> 0 Then
            ReadIDCardNoSS = "-2^二代证寻卡上电失败，错误代码：" & CStr(iReturn)
            ss_reader_close (handle)
            Exit Function
        Else
            iReturn = ss_id_read_card(handle, 0)
            If iReturn <> 0 Then
                ReadIDCardNoSS = "-3^二代证读卡失败，错误代码：" & CStr(iReturn)
                ss_reader_close (handle)
                Exit Function
            Else
                ReadStatus = 1
            End If
        End If
        ReadNum = ReadNum - 1
        If ReadStatus = 0 Then
            Sleep (1000)
        End If
'    Loop
    If ReadStatus = 0 Then
        ss_reader_close (handle)
        Exit Function
    End If
        
        
    iReturn = ss_id_query_number(handle, IDCard)
    ss_reader_close (handle)
    
    IDCard = Replace(IDCard, Chr(0), "")
    ReadIDCardNoSS = "0^" & IDCard & "^" & "" & "^"
    
    Exit Function
End Function
'读社保卡
Public Function ReadInusCardNo(CardType As String) As String

    Dim Rtn As Integer
    Dim m1str As String
    Dim Person As String
   Call InitInfo
   Device = "SS"
   If (Device = "DK") Then
        m1str = ReadInusCardNoDK(CardType)
    Else
        m1str = ReadInusCardNoSS(CardType)
    End If
    ReadInusCardNo = m1str

End Function

Public Function ReadInusCardNoDK(CardType As String) As String
    DownloadRegistration
    Dim lRet As Integer
    Dim Rtn As Integer
    Dim iType As Integer
    Dim pchOutInfo As String * 20480
    
    
    Dim pErrMsg(64) As Byte
    Rtn = DeviceBeep(10, pErrMsg(0))
    If Rtn <> 0 Then
        ReadInusCardNoDK = "-1^" & "设备蜂鸣失败"
        Exit Function
    End If
    iType = iCheckCardType()
    If (iType = 2) Or (iType = 5) Or (iType = 6) Then
        ReadInusCardNoDK = "-100^卡类型错误"
        Exit Function
    End If
    
    Dim KSBM(256) As Byte
    Dim KBL(256) As Byte
    Dim GFBB(256) As Byte
    Dim JGBM(256) As Byte
    Dim FKRQ(256) As Byte
    Dim KYZQ(256) As Byte
    Dim KH(256) As Byte

    lRet = iReadCardPublicInfo(KSBM(0), KBL(0), GFBB(0), JGBM(0), FKRQ(0), KYZQ(0), KH(0), pErrMsg(0))
    If (lRet < 0) Then
        ReadInusCardNoDK = lRet & "^" & StrConv(pErrMsg(), vbUnicode)
        Exit Function
     End If
     CardNo = CheckNullChar(StrConv(KH(), vbUnicode))
     ReadInusCardNoDK = "0" & "^" & CardNo & "^" & "" & "^"
     Exit Function

End Function

Public Function ReadInusCardNoSS(CardType As String) As String
    DownloadRegistration
    Dim hHandle As Long
    Dim lResult As Long

    hHandle = ss_reader_open()
    If (hHandle < 0) Then
        ReadInusCardNoSS = " 打开失败,错误代码：" & hHandle
        ss_reader_close (handle)
        Exit Function
    End If

    Dim no_psam As Long
    no_psam = 0
    lResult = ss_rf_sb_FindCard(no_psam)
    If (lResult <> 0) Then
        ReadInusCardNoSS = "寻卡失败,错误代码： " & lResult
        ss_reader_close (handle)
        Exit Function
    End If
    
    Dim CardIdentifier As String
    Dim CardVersion As String
    Dim IssuersID As String
    Dim IssuingDate As String
    Dim EffectiveData As String
    Dim InsuCardNo As String
    CardIdentifier = Space(1000)
    CardType = Space(1000)
    CardVersion = Space(1000)
    IssuersID = Space(1000)
    IssuingDate = Space(1000)
    EffectiveData = Space(1000)
    InsuCardNo = Space(1000)
    Ret = ss_rf_sb_ReadCardIssuers(CardIdentifier, CardType, CardVersion, IssuersID, IssuingDate, EffectiveData, InsuCardNo)
    If Ret <> 0 Then
        ReadInusCardNoSS = "社保卡读取卡号失败！" & Ret
        ss_reader_close (handle)
        Exit Function
    End If
    ss_reader_close (handle)
    InsuCardNo = CheckNullChar(InsuCardNo)
    ReadInusCardNoSS = "0^" & InsuCardNo & "^" & "" & "^"
    Exit Function

End Function
'读身份证信息
Public Function ReadIDCardInfo(CardType As String) As String
    Dim Rtn As Integer
    Dim m1str As String
    Dim Person As String
   Call InitInfo
   Device = "SS"
   'MsgBox "新版本测试...."
   If (Device = "DK") Then
        m1str = ReadIDCardInfoDK(CardType)
    Else
        m1str = ReadIDCardInfoSS(CardType)
        
    End If
    ReadIDCardInfo = m1str


End Function
Public Function ReadIDCardInfoDK(CardType As String) As String
    DownloadRegistration
    Dim lRet As Integer
    Dim TRet As Integer
    Dim iType As Integer
    Dim PatNme As String
    Dim PatSexCode As String
    Dim National As String
    Dim Brith As String
    Dim address As String
    Dim iDNo As String
    Dim age As String
    Dim InsuCardNo As String
    Dim pchOutInfo As String * 20480
    Dim pchInsuOutInfo As String * 2048
    Dim PatArr() As String
    Dim myXMLData As String
    Dim Rtn As Integer
    Dim CName As String
    Dim pErrMsg(64) As Byte
    Rtn = DeviceBeep(10, pErrMsg(0))
    If Rtn <> 0 Then
        ReadIDCardInfoDK = "-1^" & "设备蜂鸣失败"
        Exit Function
    End If
    '返回1表示社保卡；返回2表示居民健康卡，返回3表示M1卡， 返回4表示二代证，返回5表示银行卡，返回6表示无卡
    iType = iCheckCardType()
    If (iType = 2) Or (iType = 3) Or (iType = 5) Or (iType = 6) Then
        ReadIDCardInfoDK = "-100^卡类型错误"
        Exit Function
    End If
    '0-自动匹配; 1-身份证; 2-社保卡; 3-健康卡; 4-外国人居留证
    lRet = iReadIdentityCard(0, pchOutInfo, 2)
    If (lRet < 0) Then
        ReadIDCardInfoDK = lRet & "^" & pchOutInfo
        Exit Function
    End If
    pchOutInfo = StrConv(pchOutInfo, vbNarrow)
    'MsgBox pchOutInfo
    PatArr = Split(pchOutInfo, "|")
    '姓名|性别|民族|出生日期|地址|身份证号|发卡日期|有效日期|签发机关|照片BASE64|
    '王楠|男|汉族|19870831|辽宁省锦州市太和区凌川路富锦家园15-18号|210703198708312038|20100105|20200105|锦州市公安局太和分局||
    If (iType = 4) Then
    
        iDNo = PatArr(5)
        '9代表新版外国人永久居留证
        If (Mid(iDNo, 1, 1)) = 9 Then
          '证件样本证件样本证件样本证件样|女|卡密码已锁定(1604)|19800101|ZHENGJIAN, YANGBEN ZHENG JIAN YANG|931586198001010028|20230808|20330807|
          PatNme = PatArr(4)
          PatSexCode = PatArr(1)
          National = ""
          Brith = PatArr(3)
          Brith = Mid(Brith, 1, 4) & "-" & Mid(Brith, 5, 2) & "-" & Mid(Brith, 7, 2)
          address = ""
          age = GetAge(iDNo)
          CName = PatArr(0)
        ElseIf Not IsNumeric((Mid(iDNo, 1, 1))) Then
         '字母代表旧版外国人永久居留证
         '证件样本证件样本证件样本证件|ZHENGJIAN,YANGBEN ZHENG JIAN YANG BEN ZHENGJIAN Y|女|PAK|19800101|PAK310080010103|20230808|20330807|1500|01|I||
          PatNme = PatArr(1)
          PatSexCode = PatArr(2)
          National = PatArr(3)
          Brith = PatArr(4)
          Brith = Mid(Brith, 1, 4) & "-" & Mid(Brith, 5, 2) & "-" & Mid(Brith, 7, 2)
          address = ""
          age = GetAgeNew(PatArr(4))
          CName = PatArr(0)
   
        Else
         '炎黄子弟
           PatNme = PatArr(0)
           PatSexCode = PatArr(1)
           National = PatArr(2)
           Brith = PatArr(3)
           Brith = Mid(Brith, 1, 4) & "-" & Mid(Brith, 5, 2) & "-" & Mid(Brith, 7, 2)
           address = PatArr(4)
           iDNo = PatArr(5)
           age = GetAge(iDNo)
        End If
    Else
        PatNme = PatArr(0)
        PatSexCode = PatArr(1)
        National = PatArr(2)
        Brith = PatArr(3)
        Brith = Mid(Brith, 1, 4) & "-" & Mid(Brith, 5, 2) & "-" & Mid(Brith, 7, 2)
        address = PatArr(4)
        iDNo = PatArr(5)
        age = GetAge(iDNo)
    End If
    PatNme = Trim(PatNme)
    PatSexCode = getSexCode(PatSexCode)
    National = getNationCode(National)
    address = Trim(TrimASCII(address))
    iDNo = Trim(TrimASCII(iDNo))
    
    myXMLData = myXMLData & GetXMLNodeData(gPeopleName, PatNme)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleSex, PatSexCode)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleNation, National)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleBirthday, Brith)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleAddress, address)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleIDCode, iDNo)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleAge, age)
    myXMLData = myXMLData & GetXMLNodeData(gInsuCardNo, InsuCardNo)
    myXMLData = myXMLData & GetXMLNodeData(gCName, CName)
    myXMLData = "<" & gRoot & ">" & myXMLData & "</" & gRoot & ">"
    
    ReadIDCardInfoDK = "0" & "^" & myXMLData
    Exit Function

End Function
Public Function ReadIDCardInfoSS(CardType As String) As String
    DownloadRegistration
    Dim iReturn As Long
    Dim handle As Long
    Dim IDCard As String, name As String, sex As String, Folk As String, Brith As String, address As String, PhotoType As Integer, PhotoTypeAddress As String
    iReturn = 0
    IDCard = String$(255, Chr$(0))
    name = String$(255, Chr$(0))
    sex = String$(255, Chr$(0))
    Folk = String$(255, Chr$(0))
    Brith = String$(255, Chr$(0))
    address = String$(255, Chr$(0))
    age = String$(255, Chr$(0))
    handle = ss_reader_open()
    If handle <= 0 Then
        ReadIDCardInfoSS = "-1^设备打开失败，错误代码：" & CStr(handle)
        Exit Function
    End If
    iReturn = ss_id_ResetID2Card(handle)
    If iReturn <> 0 Then
        ReadIDCardInfoSS = "-2^二代证寻卡上电失败，错误代码：" & CStr(iReturn)
        ss_reader_close (handle)
        Exit Function
    End If
    iReturn = ss_id_read_card(handle, 0)
    If iReturn <> 0 Then
        ReadIDCardInfoSS = "-3^二代证读卡失败，错误代码：" & CStr(iReturn)
        ss_reader_close (handle)
        Exit Function
    End If
    iReturn = ss_id_query_number(handle, IDCard)
    iReturn = ss_id_query_name(handle, name)
    iReturn = ss_id_query_sex(handle, sex)
    iReturn = ss_id_query_folk(handle, Folk)
    iReturn = ss_id_query_birth(handle, Brith)
    iReturn = ss_id_query_address(handle, address)
    
    PhotoType = 1
    PhotoTypeAddress = "C:\tmp"
    iReturn = ss_id_query_photo_file(handle, PhotoType, PhotoTypeAddress)
    
    ss_reader_close (handle)
    
    IDCard = Replace(IDCard, Chr(0), "")
    name = Replace(name, Chr(0), "")
    name = Replace(name, " ", "")
    sex = Replace(sex, Chr(0), "")
    Folk = Replace(Folk, Chr(0), "")
    Brith = Replace(Brith, Chr(0), "")
    Brith = Mid(Brith, 1, 4) & "-" & Mid(Brith, 5, 2) & "-" & Mid(Brith, 7, 2)
    address = Replace(address, Chr(0), "")
    age = "" 'GetAge(IDNo)
    
    myXMLData = myXMLData & GetXMLNodeData(gPeopleName, name)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleSex, sex)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleNation, Folk)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleBirthday, Brith)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleAddress, address)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleIDCode, IDCard)
    myXMLData = myXMLData & GetXMLNodeData(gPeopleAge, "")
    myXMLData = myXMLData & GetXMLNodeData(gInsuCardNo, "")
    myXMLData = "<" & gRoot & ">" & myXMLData & "</" & gRoot & ">"
    
    
    ReadIDCardInfoSS = "0" & "^" & myXMLData
    Exit Function

End Function
'读磁条卡
Public Function ReadCTCardNo(CardType As String) As String
    Dim Rtn As Integer
    Dim m1str As String
    Dim Person As String
   Call InitInfo
   Device = "SS"
   If (Device = "DK") Then
        m1str = ReadCTCardNoDK(CardType)
    Else
        m1str = ReadCTCardNoSS(CardType)
        
    End If
    ReadCTCardNo = m1str


End Function
Public Function ReadCTCardNoDK(CardType As String) As String

ReadCTCardNoDK = "-1^光标落入卡号,在德卡读卡器上刷就诊卡即可!"
End Function
Public Function ReadCTCardNoSS(CardType As String) As String
    DownloadRegistration
    Dim handle As Long
    Dim iReturn As Long
    Dim IDCard As String, name As String, CardNo As String, Folk As String, Brith As String, address As String
    Dim ReadNum As Integer, ReadStatus As Integer
    iReturn = 0
    ReadStatus = 0
    ReadNum = 5
    name = String$(255, Chr$(0))
    CardNo = String$(255, Chr$(0))
    Folk = String$(255, Chr$(0))

     iReturn = SS_CT_ReadInfo_(name, CardNo, Folk, 8000)
     If iReturn <> 0 Then
         ReadCTCardNoSS = "-2^刷卡失败，错误代码：" & CStr(iReturn)
         Exit Function
     End If

    
    ReadCTCardNoSS = "0^" & Replace(CardNo, Chr(0), "") & "^" & "" & "^"
    
    Exit Function

End Function

Private Function InitInfo()
   ' 卡配置信息
   Dim Tline, INpos1 As String
   Dim TextLine
   Dim Temp As String
   
   If Dir("C:\ClsReadCard.ini") <> "" Then
        Temp = String(255, 0)
        Temp = "C:\ClsReadCard.ini"
        Open Temp For Input As #1
        Do While Not EOF(1)
            Line Input #1, Tline
             'FTP的数据
             If InStr(1, Tline, "CardReadingDevice=") = 1 Then
                Device = Mid(Tline, 19)
                'Device = "SS"
             End If
             
        Loop
        Close #1
   Else
        MsgBox ("配置文件不存在,默认为德卡读卡器")
        
        '先读德卡，在读德生
        Dim flag As String
        'm1str = ReadCardM1T10("23")
        'Flag = Split(m1str, "^")(0)
        flag = 0
        If flag = 0 Then
            Call WriteIniKey("ReadPBOC", "CardReadingDevice", "DK", "C:\ClsReadCard.ini")
            
        Else
            'm1str = ReadCardM1DS("23")
            flag = 0 'Split(m1str, "^")(0)
            If flag = 0 Then
                Call WriteIniKey("ReadPBOC", "CardReadingDevice", "SS", "C:\ClsReadCard.ini")
            End If
        End If
           
        
   End If
   
End Function
Public Sub WriteIniKey(strSection As String, strKey As String, strValue As String, filname As String)
    WritePrivateProfileString strSection, strKey, strValue, filname
End Sub
Public Function DownloadRegistration()
    Dim AppPath  As String
    Dim FileStartPath  As String
    Dim FileToPath As String
    Dim ToFileisExist  As String
    'Dim ToFileisExist64  As String
    Dim FileName  As String
     
    On Error Resume Next
    AppPath = App.Path
    
    Dim DllArray() As String, i As Long, DllStr As String, iLength As Long
    DllStr = Dir(AppPath & "\*.dll")
    Do Until DllStr = ""
        ReDim Preserve DllArray(i)
        DllArray(i) = DllStr
        DllStr = Dir
        i = i + 1
    Loop
    'DllStr = DllArray(0)
    iLength = i
    For i = 0 To iLength - 1
        DllStr = DllArray(i)
        FileName = "\" & DllStr
        FileStartPath = AppPath & FileName
        'fisExist = Dir(FileStartPath)
        FileToPath = "C:\Windows\System32" & FileName
'        FileToPath = "D:" & FileName
        ToFileisExist = Dir(FileToPath)
        'ToFileisExist64 = Dir("C:\Windows\SysWOW64\ReadCardVerson.dll")
        If ToFileisExist = "" And DllStr <> "" Then
            '下载复制文件到系统目录
            FileCopy FileStartPath, FileToPath
            '注册文件
            Shell "cmd /c regsvr32 /u /s " & FileToPath, vbHide
            Shell "cmd /c regsvr32 /u /s C:\Windows\SysWOW64" & FileName, vbHide
            
            ' Shell "cmd /c regsvr32 " & FileToPath, vbHide
            ' Shell "cmd /c regsvr32 C:\Windows\SysWOW64" & FileName, vbHide
        End If
    Next i
    
End Function
Public Function GetXMLNodeData(Tag As String, value As String) As String
    
    GetXMLNodeData = "<" & Tag & ">" & value & "</" & Tag & ">"
End Function
Function CheckNullChar(ByVal StrGet As String)
        'lxz 检查字符串中末尾空字符
        If InStr(StrGet, vbNullChar) > 0 Then
            StrGet = Left(StrGet, InStr(StrGet, vbNullChar) - 1)
        Else
            StrGet = StrGet
        End If
        Dim i As Integer
        Dim mydata As String
        Dim NullNum(0) As Byte
        Dim NummStr As String
        Dim LenStr As Integer
        LenStr = Len(StrGet)
        NullNum(0) = 255
        NummStr = StrConv(NullNum(), vbUnicode)
        mydata = ""
        Dim SubStr As String
        For i = 1 To LenStr
        SubStr = Mid(StrGet, i, 1)
        If SubStr <> " " And NummStr <> SubStr Then
            mydata = mydata & SubStr
        End If
        Next i
        CheckNullChar = mydata
End Function
'去掉ASCII 码的  1， 2， 3，
Public Function TrimASCII(CharStr As String) As String
    Dim mystr As String
    Dim myOutCharStr As String
    Dim myLength As Integer
    Dim i As Integer
    
    myLength = Len(CharStr)
    For i = 1 To myLength
        mystr = Mid$(CharStr, i, 1)
        If mystr <> Chr(0) And mystr <> Chr(1) And mystr <> Chr(2) And mystr <> Chr(3) And Asc(mystr) <> 63 And Asc(mystr) <> 32 Then
            myOutCharStr = myOutCharStr & mystr
        Else
            Exit For
        End If
    Next
    
    TrimASCII = myOutCharStr
End Function

Public Function GetAgeNew(strID As String) As Integer

 GetAgeNew = Int((Date - CDate(Format(Mid(strID, 1, 4) & "-" & Mid(strID, 5, 2) & "-" & Mid(strID, 7, 2), "YYYY-MM-DD"))) / 365)

End Function


Public Function GetAge(strID As String) As Integer
    Dim CodeLen As Long
    
     CodeLen = Len(strID)
       Select Case CodeLen
       Case 15
            GetAge = Int((Date - CDate(Format(Mid(strID, 7, 2) & "-" & Mid(strID, 9, 2) & "-" & Mid(strID, 11, 2), "YYYY-MM-DD"))) / 365)
       Case 18
           GetAge = Int((Date - CDate(Format(Mid(strID, 7, 4) & "-" & Mid(strID, 11, 2) & "-" & Mid(strID, 13, 2), "YYYY-MM-DD"))) / 365)
       Case Else
           GetAge = 0
       End Select
    
 
    Exit Function
End Function

' 帮助函数：将字节数组转换为字符串，并替换空字符
Public Function BytesToString(ByRef byteArray() As Byte, ByVal NullFlag As Boolean) As String
    Dim i As Long
    Dim result As String

    result = ""
    For i = LBound(byteArray) To UBound(byteArray)
        If byteArray(i) = 0 And NullFlag Then
            result = result & " " ' 将空字符替换为空格
        Else
            result = result & Chr(byteArray(i))
        End If
    Next i

    BytesToString = result
End Function
Private Function ByteArrayToString(ByRef byteArray() As Byte) As String
    Dim str As String
    str = StrConv(byteArray, vbUnicode)
    ByteArrayToString = Replace(str, Chr$(0), "")
End Function
