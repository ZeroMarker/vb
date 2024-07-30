Public Const gPeopleIDCode As String = "CredNo"
Public Const gPeopleName As String = "Name"
Public Const gPeopleSex  As String = "Sex"
Public Const gPeopleNation As String = "NationDesc"
Public Const gPeopleBirthday  As String = "Birth"
Public Const gPeopleAddress  As String = "Address"
Public Const gPeopleAge As String = "Age"
Public Const gInsuCardNo As String = "PatYBCode"
Public Const gRoot As String = "IDRoot"
Public Const gCName As String = "FChineseName"
Public Device As String

'读个人身份信息。支持卡片：二代身份证、外国人居留证、社保卡、健康卡
Declare Function dc_init Lib "dcrf32.dll" (ByVal port As Integer, ByVal type_ As Integer) As Integer
'读社保卡个人基本信息文件
Declare Function dc_beep Lib "dcrf32.dll" (ByVal handel As Integer, ByVal times As Integer) As Integer
'读社保卡发卡信息文件
Declare Function dc_GetSocialSecurityCardBaseInfo Lib "dcrf32.dll" _
    (ByVal icdev As Long, _
     ByVal type_ As Integer, _
     ByVal card_code As String, _
     ByVal card_type As String, _
     ByVal version As String, _
     ByVal init_org_number As String, _
     ByVal card_issue_date As String, _
     ByVal card_expire_day As String, _
     ByVal card_number As String, _
     ByVal social_security_number As String, _
     ByVal name As String, _
     ByVal name_ex As String, _
     ByVal sex As String, _
     ByVal nation As String, _
     ByVal birth_place As String, _
     ByVal birth_day As String) As Integer
'写M1卡
Declare Function dc_exit Lib "dcrf32.dll" _
    (ByVal port As Integer) As Integer

Declare Function dc_SamAReadCardInfo Lib "dcrf32.dll" _
    (ByVal handle As Integer, _
    ByVal type_ As Integer, _
    ByRef MsgLen As Long, _
    ByRef pMsg As Byte, _
    ByRef PhotoLen As Long, _
    ByRef pPhoto As Byte, _
    ByRef FingerLen As Long, _
    ByRef pFinger As Byte, _
    ByRef ExtraLen As Long, _
    ByRef pExtra As Byte) As Integer

Declare Function dc_ParsePhotoInfo Lib "dcrf32.dll" _
    (ByVal handle As Integer, _
    ByVal type_ As Integer, _
    ByVal PhotoLen As Integer, _
    ByRef pPhoto As String, _
    ByRef Base64Len As Integer, _
    ByRef pBase64 As String) As Integer

Declare Function iReadCardBas Lib "SSCardDriver.dll" _
    (ByVal type_ As Integer, _
    ByVal pMsg As String) As Integer

Declare Function dc_ParseTextInfo Lib "dcrf32.dll" _
    (ByVal handle As Integer, _
    ByVal charset As Integer, _
    ByVal info_len As Long, _
    ByRef info As Byte, _
    ByRef name As Byte, _
    ByRef sex As Byte, _
    ByRef nation As Byte, _
    ByRef birth_day As Byte, _
    ByRef address As Byte, _
    ByRef id_number As Byte, _
    ByRef department As Byte, _
    ByRef expire_start_day As Byte, _
    ByRef expire_end_day As Byte, _
    ByRef reserved As Byte) As Integer

Declare Function dc_ParseTextInfoForForeigner Lib "dcrf32.dll" _
    (ByVal handle As Integer, _
    ByVal charset As Integer, _
    ByVal info_len As Integer, _
    ByRef info As String, _
    ByRef english_name As String, _
    ByRef sex As String, _
    ByRef id_number As String, _
    ByRef citizenship As String, _
    ByRef chinese_name As String, _
    ByRef expire_start_day As String, _
    ByRef expire_end_day As String, _
    ByRef birth_day As String, _
    ByRef version_number As String, _
    ByRef department_code As String, _
    ByRef type_sign As String, _
    ByRef reserved As String) As Integer

Declare Function dc_ParseTextInfoForNewForeigner Lib "dcrf32.dll" _
    (ByVal handle As Integer, _
    ByVal charset As Integer, _
    ByVal info_len As Integer, _
    ByRef info As String, _
    ByRef chinese_name As String, _
    ByRef sex As String, _
    ByRef renew_count As String, _
    ByRef birth_day As String, _
    ByRef english_name As String, _
    ByRef id_number As String, _
    ByRef reserved As String, _
    ByRef expire_start_day As String, _
    ByRef expire_end_day As String, _
    ByRef english_name_ex As String, _
    ByRef citizenship As String, _
    ByRef type_sign As String, _
    ByRef prev_related_info As String, _
    ByRef old_id_number As String) As Integer

Declare Function dc_ParseOtherInfo Lib "dcrf32.dll" _
    (ByVal icdev As Integer, _
    ByVal flag As Integer, _
    ByRef in_info As String, _
    ByRef out_info As String) As Integer

Declare Function dc_GetIdCardType Lib "dcrf32.dll" _
    (ByVal icdev As Integer, _
    ByVal info_len As Integer, _
    ByRef in_info() As Byte) As Integer

'通用方法-----------------
Public Declare Function ss_reader_open Lib "SS728M05_SDK.dll" () As Long
Public Declare Function ss_reader_close Lib "SS728M05_SDK.dll" (ByVal icdev As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function MessageBoxTimeout Lib "user32" Alias "MessageBoxTimeoutA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long, ByVal wlange As Long, ByVal dwTimeout As Long) As Long

'读卡器提示-----------------嗡鸣器，指示灯
Public Declare Function ss_dev_beep Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal Amount As String, ByVal Msec As String) As Long
Public Declare Function ss_dev_led Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal LedClr As Byte, ByVal LedCtrl As Byte, ByVal Amount As String, ByVal Msec As String) As Long

'监听键盘事件-------------------------------------
Public Declare Function GetKeyboardInput Lib "KeyboardHook.dll" (ByVal KeyData As String, ByVal TimeOut As Long) As Long
Public Declare Function SS_CT_ReadInfo_ Lib "SS728M05_SDK.dll" (ByVal buf1 As String, ByVal buf2 As String, ByVal buf3 As String, TimeOut As Long) As Long
Public Declare Function SS_MC_SetMode Lib "SS728M05_SDK.dll" (ByVal icdev As Integer, ByVal Mode As Integer, ByVal WaitTime As Integer) As Integer
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'读社保卡----------------
'寻卡
Public Declare Function ss_rf_sb_FindCard Lib "SS728M05_SDK.dll" (ByVal no_psam As Integer) As Long
'读取发卡机构信息
Public Declare Function ss_rf_sb_ReadCardIssuers Lib "SS728M05_SDK.dll" (ByVal CardIdentifier As String, ByVal CardType As String, ByVal CardVersion As String, ByVal IssuersID As String, ByVal IssuingDate As String, ByVal EffectiveData As String, ByVal CardID As String) As Long
'读取持卡人信息
Public Declare Function ss_rf_sb_ReadCardholder Lib "SS728M05_SDK.dll" (ByVal CardID As String, ByVal name As String, ByVal name_ As String, ByVal sex As String, ByVal nation As String, ByVal address As String, ByVal Birthday As String) As Long

'读居民身份号-------------
'推荐读卡流程:打开设备'卡片复位'读卡'获得数据'关闭设备
'复位
Public Declare Function ss_id_ResetID2Card Lib "SS728M05_SDK.dll" (ByVal icdev As Long) As Long
'读卡
Public Declare Function ss_id_read_card Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal flag As Long) As Long
'获取数据
Public Declare Function ss_id_query_number Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal num As String) As Long
Public Declare Function ss_id_query_name Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal num As String) As Long
Public Declare Function ss_id_query_sex Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal num As String) As Long
Public Declare Function ss_id_query_folkL Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal num As String) As Long
Public Declare Function ss_id_query_folk Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal num As String) As Long
Public Declare Function ss_id_query_birth Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal num As String) As Long
Public Declare Function ss_id_query_address Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal num As String) As Long
Public Declare Function ss_id_query_photo_file Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal Format As Integer, ByVal ImagePath As String) As Long


'M1卡操作-------------------
'打开设备'卡片复位'校验秘钥'读写操作'关闭设备
'卡片复位
Public Declare Function ss_CardMifare_Reset Lib "SS728M05_SDK.dll" (ByVal icdev As Long) As Long
'校验秘钥
'Public Declare Function ss_CardMifare_Authentication Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal Mode As String, ByVal SecNr As String, ByVal Password As String) As Long
Public Declare Function ss_CardMifare_Authentication Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal Mode As Byte, ByVal SecNr As Byte, ByRef Password As Byte) As Long
'读取数据块内数据
Public Declare Function ss_CardMifare_ReadBlock Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal Adr As Byte, ByRef Data As Byte) As Long
'向数据块内写数据
Public Declare Function ss_CardMifare_WriteBlock Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal Adr As String, ByVal Data As String) As Long
'读取M1卡UID
Public Declare Function ss_CardMifare_GetUID Lib "SS728M05_SDK.dll" (ByVal icdev As Long, ByVal UID As String) As Long

'读取M1卡UID
Public Declare Function GetKeyA Lib "VCEncrypt.dll" (ByVal strUID As String, ByVal strKeyA As String) As Long
'校验数据
Public Declare Function CheckData Lib "VCEncrypt.dll" (ByVal strUID As String, ByVal strData As String) As Long

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'''
'[DllImport("dcrf32.dll")]
'public static extern int dc_init(int port, int type);
'[DllImport("dcrf32.dll")]
'public static extern int dc_beep(int handel, int times);
'[DllImport("dcrf32.dll")]
'public static extern int dc_SamAReadCardInfo(int handel, int type, out int msgLen, ref byte pMsg, out int photoLen, ref byte pPhoto, out int fingerLen, ref byte pFinger, out int extraLen, ref byte pExtra);
'[DllImport("dcrf32.dll")]
'//public static extern int dc_ParsePhotoInfo(int handel, int type, int photoLen, ref byte pPhoto, out int base64Len, ref byte pBase64);
'public static extern int dc_ParsePhotoInfo(int handel, int type, int photoLen, ref byte pPhoto, out int base64Len, out byte pBase64);
'[DllImport("dcrf32.dll")]
'public static extern int dc_exit(int port);
'[DllImport("SSCardDriver.dll")]
'public static extern int iReadCardBas(int type, StringBuilder pMsg);

'[DllImport("dcrf32.dll")]
'public static extern short dc_ParseTextInfo(int handle, int charset, int info_len, [Out] byte[] info, [Out] byte[] name, [Out] byte[] sex, [Out] byte[] nation, [Out] byte[] birth_day, [Out] byte[] address, [Out] byte[] id_number, [Out] byte[] department, [Out] byte[] expire_start_day, [Out] byte[] expire_end_day, [Out] byte[] reserved);
'[DllImport("dcrf32.dll")]
'public static extern short dc_ParseTextInfoForForeigner(int handle, int charset, int info_len, [Out] byte[] info, [Out] byte[] english_name, [Out] byte[] sex, [Out] byte[] id_number, [Out] byte[] citizenship, [Out] byte[] chinese_name, [Out] byte[] expire_start_day, [Out] byte[] expire_end_day, [Out] byte[] birth_day, [Out] byte[] version_number, [Out] byte[] department_code, [Out] byte[] type_sign, [Out] byte[] reserved);

'[DllImport("dcrf32.dll")]
'public static extern short dc_ParseTextInfoForNewForeigner(int handle, int charset, int info_len, [Out] byte[] info, [Out] byte[] chinese_name, [Out] byte[] sex, [Out] byte[] renew_count, [Out] byte[] birth_day, [Out] byte[] english_name, [Out] byte[] id_number, [Out] byte[] reserved, [Out] byte[] expire_start_day, [Out] byte[] expire_end_day, [Out] byte[] english_name_ex, [Out] byte[] citizenship, [Out] byte[] type_sign, [Out] byte[] prev_related_info, [Out] byte[] old_id_number);

'[DllImport("dcrf32.dll")]
'public static extern short dc_ParsePhotoInfo(int handle, int type, int info_len, [Out] byte[] info, ref int photo_len, [Out] byte[] photo);
'[DllImport("dcrf32.dll")]
'public static extern short dc_ParseOtherInfo(int icdev, int flag, [In] byte[] in_info, [Out] byte[] out_info);

'[DllImport("dcrf32.dll")]
'public static extern short dc_GetIdCardType(int icdev, int info_len, [In] byte[] in_info);

'[DllImport("dcrf32.dll")]
'public static extern short dc_GetSocialSecurityCardBaseInfo(int icdev, int type, [Out] byte[] card_code, [Out] byte[] card_type, [Out] byte[] version, [Out] byte[] init_org_number, [Out] byte[] card_issue_date, [Out] byte[] card_expire_day, [Out] byte[] card_number, [Out] byte[] social_security_number, [Out] byte[] name, [Out] byte[] name_ex, [Out] byte[] sex, [Out] byte[] nation, [Out] byte[] birth_place, [Out] byte[] birth_day);
'''


'读取M1卡数据
'Declare Function iReadM1Card Lib "dcrf32.dll" (ByVal SecNr%, ByVal BlockNr%, ByVal SecKey$, ByVal KeyMode%, ByVal pOutInfo$) As Integer
'获取放置的卡类型
'1:社保卡;2:居民健康卡;3:M1卡; 4:二代证;5:银行卡;6:无卡;其他值未定义
'Declare Function iCheckCardType Lib "DC_Reader.dll" () As Integer


'读个人身份信息。支持卡片：二代身份证、外国人居留证、社保卡、健康卡
Declare Function iReadIdentityCard Lib "DC_Reader.dll" (ByVal iType%, ByVal pOutInfo$, ByVal newtout%) As Integer
'读社保卡个人基本信息文件
Declare Function iReadSIEF06 Lib "DC_Reader.dll" (ByVal pchOutInfo$) As Integer
'读社保卡发卡信息文件
Declare Function iReadSIEF05 Lib "DC_Reader.dll" (ByVal pchOutInfo$) As Integer
'写M1卡
Declare Function iWriteM1Card Lib "DC_Reader.dll" (ByVal SecNr%, ByVal BlockNr%, ByVal SecKey$, ByVal KeyMode%, ByVal Databuff$, ByVal pOutInfo$) As Integer
'读取M1卡数据
Declare Function iReadM1Card Lib "DC_Reader.dll" (ByVal SecNr%, ByVal BlockNr%, ByVal SecKey$, ByVal KeyMode%, ByVal pOutInfo$) As Integer
'获取放置的卡类型
'1:社保卡;2:居民健康卡;3:M1卡; 4:二代证;5:银行卡;6:无卡;其他值未定义
Declare Function iCheckCardType Lib "DC_Reader.dll" () As Integer

Public Declare Function DeviceBeep Lib "SSCardDriver.dll" (ByVal time As Integer, ByRef pErrMsg As Byte) As Long

Public Declare Function ReadMifare Lib "SSCardDriver.dll" (ByVal KeyMode As Integer, ByVal SecNr As Integer, ByVal SecKey As String, ByVal BlockNr As Integer, ByRef SecMsg As Byte, ByRef pcErrM As Byte) As Long

Public Declare Function iReadCardPublicInfo Lib "SSCardDriver.dll" (ByRef KSBM As Byte, ByRef KBL As Byte, ByRef GFBB As Byte, ByRef JGBM As Byte, ByRef FKRQ As Byte, ByRef KYZQ As Byte, ByRef KH As Byte, ByRef pErrMsg As Byte) As Long







Public Function getSexCode(SexDesc)
    Dim SexCode As String
    SexCode = SexDesc
    If SexDesc = "男" Then
        SexCode = "1"
    ElseIf SexDesc = "女" Then
        SexCode = "2"
    End If
    
    getSexCode = SexCode
End Function


Public Function getNationCode(NationDesc)
    Dim nation As String
   nation = NationDesc
   If nation = "汉" Then
    nation = 1

   ElseIf nation = "蒙古" Then
    nation = 2
 
   ElseIf nation = "回" Then
    nation = 3
  
   ElseIf nation = "藏" Then
    nation = 4
  
   ElseIf nation = "维吾尔" Then
    nation = 5
  
   ElseIf nation = "苗" Then
    nation = 6
   
   ElseIf nation = "彝" Then
    nation = 7
  
   ElseIf nation = "壮" Then
    nation = 8
   
   ElseIf nation = "布依" Then
    nation = 9
  
   ElseIf nation = "朝鲜" Then
    nation = 10
  
   ElseIf nation = "满" Then
    nation = 11
   
   ElseIf nation = "侗" Then
    nation = 12
   
   ElseIf nation = "瑶" Then
    nation = 13
  
   ElseIf nation = "白" Then
    nation = 14
  
   ElseIf nation = "土家" Then
    nation = 15
 
   ElseIf nation = "哈尼" Then
    nation = 16
 
   ElseIf nation = "哈萨克" Then
    nation = 17
  
   ElseIf nation = "傣" Then
    nation = 18
  
   ElseIf nation = "黎" Then
    nation = 19
    
   ElseIf nation = "僳僳" Then
    nation = 20
  
   ElseIf nation = "佤" Then
    nation = 21

   ElseIf nation = "畲" Then
    nation = 22
    
   ElseIf nation = "高山" Then
    nation = 23

   ElseIf nation = "拉祜" Then
    nation = 24

   ElseIf nation = "水" Then
    nation = 25

   ElseIf nation = "东乡" Then
    nation = 26

   ElseIf nation = "纳西" Then
    nation = 27

   ElseIf nation = "景颇" Then
    nation = 28

   ElseIf nation = "柯尔克孜" Then
    nation = 29

   ElseIf nation = "土" Then
    nation = 30
  
   ElseIf nation = "达斡尔" Then
    nation = 31
    
   ElseIf nation = "仫佬" Then
    nation = 32
    
   ElseIf nation = "羌" Then
    nation = 33
  
   ElseIf nation = "布朗" Then
    nation = 34
   
   ElseIf nation = "撒拉" Then
    nation = 35
  
   ElseIf nation = "毛难" Then
    nation = 36

   ElseIf nation = "仡佬" Then
    nation = 37
    
   ElseIf nation = "锡伯" Then
    nation = 38

   ElseIf nation = "阿昌" Then
    nation = 39
 
   ElseIf nation = "普米" Then
    nation = 40
  
   ElseIf nation = "塔吉克" Then
    nation = 41

   ElseIf nation = "怒" Then
    nation = 42
   
   ElseIf nation = "乌孜别克" Then
    nation = 43

   ElseIf nation = "俄罗斯" Then
    nation = 44

   ElseIf nation = "鄂温克" Then
    nation = 45
  
   ElseIf nation = "崩龙" Then
    nation = 46

   ElseIf nation = "保安" Then
    nation = 47

   ElseIf nation = "裕固" Then
    nation = 48

   ElseIf nation = "京" Then
    nation = 49
   
   ElseIf nation = "塔塔尔" Then
    nation = 50

   ElseIf nation = "独龙" Then
    nation = 51
    
   ElseIf nation = "鄂伦春" Then
    nation = 52
    
   ElseIf nation = "赫哲" Then
    nation = 53
    
   ElseIf nation = "门巴" Then
    nation = 54
    
   ElseIf nation = "珞巴" Then
    nation = 55
    
   ElseIf nation = "基诺" Then
    nation = 56
    
   ElseIf nation = "其他" Then
    nation = 57
   End If

    getNationCode = nation
End Function


