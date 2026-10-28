Attribute VB_Name = "Common_Module"
'****************************************************************
'*                                                              *
'*  Common_Module                                               *
'*                                                              *
'*      作成日  ：2007/02/14                                    *
'*      作成者  ：澤原                                          *
'*      ｺﾒﾝﾄ    ：共通変数・関数定義                            *
'*                                                              *
'****************************************************************


'Option Explicit

'---------------------------------------------
'外部API関数使用宣言
'---------------------------------------------
'---iniﾌｧｲﾙ取得用API
Private Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
        ByVal lpDefault As String, ByVal lpReturnedString As String, _
        ByVal nSize As Long, ByVal lpFileName As String) As Long
'---iniﾌｧｲﾙ取得用API
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" _
        Alias "GetPrivateProfileSectionNamesA" _
        (ByVal lpReturnedString As String, ByVal nSize As Long, _
        ByVal lpFileName As String) As Long


'---------------------------------------------
'ｸﾞﾛｰﾊﾞﾙ変数定義
'---------------------------------------------
'---共通列挙型定義
Public Enum Ables
    DISABLE = 0
    ENABLE = 1
End Enum

'---共通変数群
Public SelectedSlot As Integer          '選択ｽﾛｯﾄ

Public WorkFolder As String             'ROMPainterがｲﾝｽﾄｰﾙされているﾊﾟｽ

Public SelectedFilePass As String       '選択されたﾌｧｲﾙのﾌﾙﾊﾟｽ
Public SaveFilePass As String           '↑を保存
Public FileSize As Long                 '選択されたﾌｧｲﾙのｻｲｽﾞ
Public AllFileSize As Long              '選択されたﾌｧｲﾙのｻｲｽﾞ(保存用)
Public FileNo As Integer                '選択されたﾌｧｲﾙのﾌｧｲﾙ番号
Public FilePointer As Variant           '選択されたﾌｧｲﾙの読出し位置
Public FileData(511) As Byte            '選択されたﾌｧｲﾙのﾃﾞｰﾀ
Public SumValue(1) As Byte              'ｻﾑ値
Public CalSumValue As Long              '計算ｻﾑ値

Public SystemAddress As String          'ｼｽﾃﾑ先頭ｱﾄﾞﾚｽ
Public SystemFilePass As String         'ｼｽﾃﾑﾌｧｲﾙのﾊﾟｽ
Public SystemFileSize As String         'ｼｽﾃﾑﾌｧｲﾙのｻｲｽﾞ

Public DevInfoAddress As String         'ﾃﾞﾊﾞｲｽﾈｯﾄ情報先頭ｱﾄﾞﾚｽ
Public DevInfoFilePass As String        'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙのﾊﾟｽ
Public DevInfoFileSize As String        'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙのｻｲｽﾞ

Public CfgInfoAddress As String         'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報先頭ｱﾄﾞﾚｽ
Public CfgInfoFilePass As String        'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙのﾊﾟｽ
Public CfgInfoFileSize As String        'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙのｻｲｽﾞ

Public PassForCancel As String          'ﾌｧｲﾙﾊﾟｽ一時退避用
Public SizeForCancel As Long            'ﾌｧｲﾙｻｲｽﾞ一時退避用

Public McomDataNum As Integer           'ﾒｿｯﾄﾞ1回のﾃﾞｰﾀｻｲｽﾞ(Mｺﾏﾝﾄﾞのﾃﾞｰﾀ部分のｻｲｽﾞ)

Public IOinfo(15) As String             'I/O構成情報

Public KpRomWindow As Integer           'FPGAのｳｨﾝﾄﾞｳ切替え
Public MXA_U_Value As Integer           'MXA_Uﾚｼﾞｽﾀに設定するﾌﾗｯｼｭのｱﾄﾞﾚｽ上位ﾋﾞｯﾄ

Public Enum Endians                     'ｴﾝﾃﾞｨｱﾝ
    BIG = 0
    LITTLE = 1
End Enum
Public Endian As Endians

Public Enum CpuNos                     'CPU番号
    CPU1 = 1
    CPU2 = 2
End Enum
Public CpuNo As CpuNos


Public Enum MODULE_TYPES                'ﾓｼﾞｭｰﾙﾀｲﾌﾟ
    NONE_TYPE = -1
    PFKP_TYPE = 0
    CMDGA_TYPE = 1
    DEVNET_TYPE = 2
    FLNET_TYPE = 3
    RTOS_TYPE = 4
    NSERIAL_TYPE = 5
    ETHER_TYPE = 6
    NETHER_TYPE = 7
    AROUSAL_TYPE = 8
End Enum
Public ModuleType As MODULE_TYPES


'---ROM消去用変数群
Public EraseAddress As Long             '消去ｱﾄﾞﾚｽ
Public EraseSize As Integer             '消去ｻｲｽﾞ
Public Enum EraseModes                  '消去ﾓｰﾄﾞ
    Tip_Erase = 0
    Sector_Erase = 1
    Cancel_Erase = 2
    Error_Erase = -1
End Enum
Public EraseMode As EraseModes

'---ROM書込み用変数群
Public WriteAddress As Long             '書込みｱﾄﾞﾚｽ
Public WriteSize As Long                '書込みｻｲｽﾞ
Public AllWriteSize As Long             '全書込みｻｲｽﾞ
Public WriteFirm As Ables               'ﾌｧｰﾑの書込みの有無
Public WriteName As Ables               'ﾓｼﾞｭｰﾙ形名書込みの有無
Public WriteModuleNameData(9) As Byte   '書込みﾓｼﾞｭｰﾙ名
Public MacWriteAddress As Long          'MACｱﾄﾞﾚｽ書込み番地
Public WriteMacData(9) As Byte          'MACｱﾄﾞﾚｽ
Public WriteMAC As Ables                'MACｱﾄﾞﾚｽ書込みの有無
Public WriteSystem As Ables             'ｼｽﾃﾑﾌﾞｰﾄ書込みの有無
Public WriteDevInfo As Ables            'ﾃﾞﾊﾞｲｽﾈｯﾄ情報書込みの有無
Public WriteCfgInfo As Ables            'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報書込みの有無
Public Enum WriteModes                  '書込みﾓｰﾄﾞ
    Normal_Write = 0
    Cancel_Write = 1
    Error_Write = -1
End Enum
Public WriteMode As WriteModes

'---ROMｺﾋﾟｰ用変数群
Public CopyAddress As Long              'ｺﾋﾟｰｱﾄﾞﾚｽ
Public CopySize As Integer              'ｺﾋﾟｰｻｲｽﾞ
Public AllCopySize As Long              '全ｺﾋﾟｰｻｲｽﾞ
Public TargetSlot As Integer            'ｺﾋﾟｰ先ｽﾛｯﾄ
Public CopyFirmData(127) As Long        '読出したﾌｧｰﾑのﾃﾞｰﾀ
Public CopyFilePass As String           '読出しﾃﾞｰﾀ保存ﾌｧｲﾙのﾊﾟｽ
Public ChangeMAC As Ables               'MACｱﾄﾞﾚｽ書換えの有無
Public Enum CopyModes                   'ｺﾋﾟｰﾓｰﾄﾞ
    Normal_Copy = 0
    Cancel_Copy = 2
    Error_Copy = -1
End Enum
Public CopyMode As CopyModes

'---ROM照合用変数群
Public VerifyAddress As Long            '照合ｱﾄﾞﾚｽ
Public VerifySize As Long               '照合ｻｲｽﾞ
Public AllVerifySize As Long            '全照合ｻｲｽﾞ
Public VerifyErrorCount As Long         '照合不一致数
Public MacWrittenAddress As Long        'MACｱﾄﾞﾚｽが書き込まれている番地
Public VerifyFirm As Ables              'F/W照合の有無
Public VerifyDevInfo As Ables           'ﾃﾞﾊﾞｲｽﾈｯﾄ情報照合の有無
Public VerifyCfgInfo As Ables           'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報照合の有無
Public VerifyMAC As Ables               'MACｱﾄﾞﾚｽ照合の有無
Public VerifySystem As Ables            'ｼｽﾃﾑ情報照合の有無
Public Enum VerifyModes                 '照合ﾓｰﾄﾞ
    Normal_Verify = 0
    Cancel_Verify = 1
    Over_Verify = 2
    Error_Verify = -1
End Enum
Public VerifyMode As VerifyModes

'---ROM読出し用変数群
Public ReadAddress As Long              '読出し開始ｱﾄﾞﾚｽ
Public ReadSize As Long                 '読出しｻｲｽﾞ
Public AllReadSize As Long              '読出しｻｲｽﾞ(保存用)
Public ReadFileName As String           '読出しﾃﾞｰﾀ保存ﾌｧｲﾙ名
Public ReadFilePass As String           '読出しﾃﾞｰﾀ保存ﾌｧｲﾙのﾊﾟｽ
Public InitMacAddress As Ables          'MACｱﾄﾞﾚｽ初期化の有無
Public MacReadAddress As Long           'MACｱﾄﾞﾚｽ読出し番地
Public DefaultMAC(4) As Integer         'ﾃﾞﾌｫﾙﾄMACｱﾄﾞﾚｽ
Public Enum ReadModes                   '読出しﾓｰﾄﾞ
    Normal_Read = 0
    Overwrite_Read = 1
    Cancel_Read = 2
    Error_Read = -1
End Enum
Public ReadMode As ReadModes


'---------------------------------------------
'ComMng_Moduleｸﾗｽのｵﾌﾞｼﾞｪｸﾄ変数を用意
'---------------------------------------------
Public cmm As ComMng_Module


'---------------------------------------------
'名前定義
'---------------------------------------------
'---CMDGAのﾚｼﾞｽﾀ
Public Const ZIOSEL3 As Long = &HBF00
'---KPGAのﾚｼﾞｽﾀNo（～200）
Public Const SYS_CTL As Long = &H0
Public Const PORT0_CTL As Long = &HA0
Public Const PORT0_STS As Long = &HA2
Public Const PORT1_CTL As Long = &HA4
Public Const PORT1_STS As Long = &HA6
'---KPGAのﾚｼﾞｽﾀNo（200～）
Public Const WDT_CTL As Long = &H14
Public Const SLAVE_CMD As Long = &H70
'---PFGAのﾚｼﾞｽﾀ
Public Const SYS_CTL_PFKP As Long = &H60B001
Public Const PORT0_MODE As Long = &H60B095
Public Const PORT0_CTLSTS As Long = &H60B097
'---FPGAのﾚｼﾞｽﾀ
Public Const ROM_WRITE_ADDRESS = &H2B64
Public Const ROM_WINDOW_ADDRESS = &H2B66
'---MXKPNのﾚｼﾞｽﾀ
Public Const TRISTATE_REG = &H200010
Public Const MXA_U_REG = &H200012

'---定数
Public Const BYTE_UNIT As Integer = 1
Public Const WORD_UNIT As Integer = 2
Public Const LWORD_UNIT As Integer = 4
Public Const OK As Integer = 0
Public Const NG As Integer = -1

'---RTOS CPUのｱﾄﾞﾚｽ
Public Const RTOS_MODULENAME_ADDRESS As Long = &HDFF0   'ﾓｼﾞｭｰﾙ形名を書き込むｱﾄﾞﾚｽ
Public Const RTOS_MACADDRESS_ADDRESS As Long = &HFF00   'MACｱﾄﾞﾚｽを書き込むｱﾄﾞﾚｽ

'---FlashROMのｱﾄﾞﾚｽ
Public Const PL64j_SA6 As Long = &HC000
Public Const PL64j_SA7 As Long = &HE000



'---------------------------------------------------------------
'
'   FAM3Connect()
'
'       機能    ：FAM3と接続
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/14 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub FAM3Connect()
    
    Dim i As Integer
    Dim s_PortNum As String
    
    '通信設定
    If Mid$(Main_Form.ConnectTarget_cmb.Text, 1, 3) = "COM" Then
        For i = 1 To Len(Main_Form.ConnectTarget_cmb.Text)
            Select Case Asc(Mid(Main_Form.ConnectTarget_cmb.Text, i, 1))
            Case 48 To 57
                s_PortNum = s_PortNum & Mid(Main_Form.ConnectTarget_cmb.Text, i, 1)
            Case Else
                '数値以外の場合何もしない
            End Select
        Next
        cmm.ComPort = Int(s_PortNum)
        cmm.ConnectMethod = Serial
        cmm.TargetCpuNo = CpuNo
    ElseIf Mid$(Main_Form.ConnectTarget_cmb.Text, 1, 3) = "USB" Then
        cmm.ConnectMethod = USB
        MsgBox "USBには未対応です", vbExclamation, "Sorry!"
        Exit Sub
    Else
        cmm.ConnectMethod = ETHER
        cmm.IPAddress = Main_Form.ConnectTarget_cmb.Text
        cmm.TargetCpuNo = CpuNo
    End If
    
    'FAM3と接続
    Call cmm.Connect
    If cmm.ConnectState = Connecting Then
        Call CheckTargetList
        If cmm.ConnectMethod = Serial Then
            Main_Form.Line1.BorderColor = RGB(212, 172, 220)
            Main_Form.Line2.BorderColor = RGB(212, 172, 220)
        ElseIf cmm.ConnectMethod = ETHER Then
            Main_Form.Line1.BorderColor = RGB(255, 228, 67)
            Main_Form.Line2.BorderColor = RGB(255, 228, 67)
        Else
            Main_Form.Line1.BorderColor = &H80000000
            Main_Form.Line2.BorderColor = &H80000000
        End If
    ElseIf cmm.ConnectState = ConnectErr Then
        MsgBox "接続に失敗しました", vbCritical, "ERROR"
        AddSysLog ("接続に失敗しました")
        Main_Form.Line1.BorderColor = &H80000000
        Main_Form.Line2.BorderColor = &H80000000
    End If

End Sub



'---------------------------------------------------------------
'
'   FAM3DisConnect()
'
'       機能    ：FAM3から切断
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/14 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub FAM3DisConnect()

    'FAM3から切断
    Call cmm.DisConnect
    If cmm.ConnectState = ConnectErr Then
        MsgBox ("切断に失敗しました")
        AddSysLog ("切断に失敗しました")
    ElseIf cmm.ConnectState = NotConnect Then
        Main_Form.Line1.BorderColor = &H80000000
        Main_Form.Line2.BorderColor = &H80000000
    End If


End Sub



'---------------------------------------------------------------
'
'   AddSysLog()
'
'       機能    ：ｼｽﾃﾑﾛｸﾞにﾛｸﾞ追加
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/15 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub AddSysLog(AddStrings As String)

    Main_Form.SysLog_txt.SelStart = Len(Main_Form.SysLog_txt.Text)
    Main_Form.SysLog_txt.SelText = AddStrings + vbCrLf

End Sub



'---------------------------------------------------------------
'
'   CheckTargetList()
'
'       機能    ：接続先にﾃｷｽﾄ入力されたものがﾘｽﾄ内にあるかﾁｪｯｸ
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/20 by 澤原
'       ｺﾒﾝﾄ    ：なければﾘｽﾄに追加
'
'---------------------------------------------------------------
Public Sub CheckTargetList()

    Dim tmpAddress As String, i As Integer, ExistFlag As Integer
    
    If Main_Form.ConnectTarget_cmb.ListIndex < 0 Then            'ﾘｽﾄ部分の項目の選択ではなくﾃｷｽﾄ部分への直接入力 = -1
        ExistFlag = 0
        tmpAddress = Main_Form.ConnectTarget_cmb.Text
        For i = 0 To Main_Form.ConnectTarget_cmb.ListCount - 1   'ﾘｽﾄ部分に存在するかﾁｪｯｸ
            If Main_Form.ConnectTarget_cmb.List(i) = tmpAddress Then
                ExistFlag = 1
            End If
        Next
        If ExistFlag = 0 Then
            Main_Form.ConnectTarget_cmb.AddItem (tmpAddress)     'ﾘｽﾄ部分になければ項目追加
        End If
    End If

End Sub



'---------------------------------------------------------------
'
'   GetINIValue()
'
'       機能    ：iniﾌｧｲﾙ取得
'       引数    ：
'       戻り値  ：iniﾌｧｲﾙのｷｰの値
'       作成    ：2007/02/20 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Function GetINIValue(KEY As String, Section As String, ININame As String) As String

    Dim Value As String * 255
    
    Call GetPrivateProfileString(Section, KEY, "ERROR", Value, 255, ININame)
    '戻り値
    GetINIValue = Left$(Value, InStr(1, Value, vbNullChar) - 1)

End Function



'---------------------------------------------------------------
'
'   CheckModuleName()
'
'       機能    ：ﾓｼﾞｭｰﾙ名ﾁｪｯｸ
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/20 by 澤原
'       ｺﾒﾝﾄ    ：①ﾓｼﾞｭｰﾙ名がiniﾌｧｲﾙ内にあるかﾁｪｯｸ
'                 ②あれば処理方式を指定(ModuleType変数の設定)
'
'---------------------------------------------------------------
Public Function CheckModuleName() As Integer

    Dim ModuleNameKey As String
    Dim SerchStr As String
    Dim i As Integer
    Dim res As String
    
    Dim strGetBuffer As String
    Dim strLastNo As Integer
    Dim LastSecName As String
    Dim LastSecNo As Integer
    Dim NullPos As Integer
        
    
    'iniﾌｧｲﾙ内の最後のｾｸｼｮﾝ名を取得
    strGetBuffer = String(100, vbNullChar)
    Call GetPrivateProfileSectionNames(strGetBuffer, Len(strGetBuffer), App.Path & "\ModuleName.ini")   '全てのｾｸｼｮﾝ名をﾊﾞｯﾌｧに格納
    strLastNo = InStr(strGetBuffer, vbNullChar + vbNullChar)    'ﾊﾞｯﾌｧ内でNullが二回続く位置
    LastSecName = Mid(strGetBuffer, strLastNo - 2, 2)           '一番後ろのｾｸｼｮﾝ名(String)
    NullPos = InStr(LastSecName, vbNullChar)                    'Nullの位置を探す
    If NullPos Then                                             'Nullの位置が先頭なら、（ｾｸｼｮﾝ名が一桁）
        LastSecName = Mid(LastSecName, 2)                       '二文字目から取り出す
    End If
    LastSecNo = CInt(LastSecName)                               '一番後ろのｾｸｼｮﾝ名(Integer型)
    
    
    'iniﾌｧｲﾙからﾓｼﾞｭｰﾙ名が対応しているか探す
    SerchStr = Main_Form.ModuleName(SelectedSlot - 1).Text     '検索するﾓｼﾞｭｰﾙ名
    For i = 0 To LastSecNo
        ModuleNameKey = GetINIValue("MODULENAME", Str(i), App.Path & "\ModuleName.ini") '検索元
        res = InStr(ModuleNameKey, SerchStr)    'ｷｰから検索
        If res >= 1 Then
            ModuleType = i      'ﾓｼﾞｭｰﾙ名があったら、その方式をｾｯﾄ
            Exit Function
        End If
    Next
    
    ModuleType = NONE_TYPE

End Function



'---------------------------------------------------------------
'
'   ErrorCheck()
'
'       機能    ：ﾓｼﾞｭｰﾙ名指定ﾁｪｯｸ、対応ﾓｼﾞｭｰﾙﾁｪｯｸ、ｽﾛｯﾄ3,4CPUﾁｪｯｸ
'       引数    ：0ﾁｪｯｸ無、1ﾁｪｯｸ有
'       戻り値  ：
'       作成    ：2007/03/05 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Function ErrorCheck(s_NameCheck As Integer, s_IniCheck As Integer, _
                            s_Slot34CpuCheck As Integer, s_FileCheck As Integer) As Integer

    Dim i As Integer

    'ﾓｼﾞｭｰﾙ名指定ﾁｪｯｸ
    If s_NameCheck Then
        If Main_Form.ModuleName(SelectedSlot - 1).Text = "" Then
            MsgBox "対象モジュールを指定して下さい", vbExclamation, "ERROR"
            ErrorCheck = -1
            Exit Function
        End If
    End If

    '対応ﾓｼﾞｭｰﾙﾁｪｯｸ
    If s_IniCheck Then
        Call CheckModuleName
        If ModuleType = NONE_TYPE Then
            MsgBox "このモジュールには対応していません", vbExclamation, "SORRY!"
            ErrorCheck = -1
            Exit Function
        End If
    End If

    'ｽﾛｯﾄ3,4 CPUﾁｪｯｸ
    If s_Slot34CpuCheck Then
        For i = 2 To 3
            If Left$(Main_Form.ModuleName(i).Text, 2) = "RP" Or Left$(Main_Form.ModuleName(i).Text, 2) = "SP" Then
                MsgBox "スロット3,4にCPU系モジュールがある場合は処理できません", vbExclamation, "SORRY!"
                ErrorCheck = -1
                Exit Function
            End If
        Next
    End If

    '選択ﾌｧｲﾙﾁｪｯｸ
    If (ModuleType = DEVNET_TYPE) Or (ModuleType = RTOS_TYPE) Then Exit Function
    If s_FileCheck Then
        If Main_Form.FileName_txt.Text = "" Then
            MsgBox "ファイルを選択して下さい", vbExclamation, "ERROR"
            ErrorCheck = -1
            Exit Function
        End If
    End If
  
End Function



'---------------------------------------------------------------
'
'   ByteToInt()
'
'       機能    ：2つのByte型⇒Integer型結合
'       引数    ：下位ﾊﾞｲﾄ, 上位ﾊﾞｲﾄ
'       戻り値  ：結合されたInteger型ﾃﾞｰﾀ
'       作成    ：2007/03/06 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Function ByteToInt(s_High As Byte, s_Low As Byte) As Integer

    If s_High > &H7F Then
        ByteToInt = s_Low + (CLng(s_High) * &H100) - &H10000
    Else
        ByteToInt = s_Low + (s_High * &H100)
    End If

End Function



'---------------------------------------------------------------
'
'   IntToLng()
'
'       機能    ：2つのInteger型⇒Long型結合
'       引数    ：下位ﾜｰﾄﾞ, 上位ﾜｰﾄﾞ
'       戻り値  ：結合されたLong型ﾃﾞｰﾀ
'       作成    ：2007/03/06 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Function IntToLng(s_High As Integer, s_Low As Integer) As Long

    IntToLng = (s_Low And &HFFFF&) Or (s_High * &H10000)

End Function



'---------------------------------------------------------------
'
'   ByteToLng()
'
'       機能    ：4つのByte型⇒Long型結合
'       引数    ：HighHighﾊﾞｲﾄ, HighLowﾊﾞｲﾄ, LowHighﾊﾞｲﾄ, LowLowﾊﾞｲﾄ
'       戻り値  ：結合されたLong型ﾃﾞｰﾀ
'       作成    ：2007/03/06 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Function ByteToLng(s_HighHigh As Byte, s_HighLow As Byte, s_LowHigh As Byte, s_LowLow As Byte) As Long

    Dim HighWord As Integer
    Dim LowWord As Integer
        
    '上位ﾜｰﾄﾞ作成
    HighWord = ByteToInt(s_HighHigh, s_HighLow)

    '下位ﾜｰﾄﾞ作成
    LowWord = ByteToInt(s_LowHigh, s_LowLow)
 
    '上位ﾜｰﾄﾞと下位ﾜｰﾄﾞ結合
    ByteToLng = IntToLng(HighWord, LowWord)
    
End Function



'---------------------------------------------------------------
'
'   IntToByte()
'
'       機能    ：Integer型⇒2つのByte型
'       引数    ：分割するInteger型ﾃﾞｰﾀ, 上位ﾊﾞｲﾄ, 下位ﾊﾞｲﾄ
'       戻り値  ：分割されたByte型ﾃﾞｰﾀ
'       作成    ：2007/03/06 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub IntToByte(s_Value As Integer, s_High As Byte, s_Low As Byte)

    '下位ﾊﾞｲﾄを切り出す
    s_Low = s_Value And &HFF

    '上位ﾊﾞｲﾄを切り出す
    If s_Value < 0 Then
        s_High = (s_Value + &H10000) \ &H100
    Else
        s_High = s_Value \ &H100
    End If

End Sub



'---------------------------------------------------------------
'
'   LngToInt()
'
'       機能    ：Long型⇒2つのInteger型
'       引数    ：分割するLong型ﾃﾞｰﾀ, 上位ﾜｰﾄﾞ, 下位ﾜｰﾄﾞ
'       戻り値  ：分割されたInteger型ﾃﾞｰﾀ
'       作成    ：2007/03/06 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub LngToInt(s_Value As Long, s_High As Integer, s_Low As Integer)

    '下位ﾜｰﾄﾞを切り出す
    If (s_Value And &HFFFF&) > &H7FFF Then
        s_Low = (s_Value And &HFFFF&) - &H10000
    Else
        s_Low = s_Value And &HFFFF&
    End If
    
    '上位ﾜｰﾄﾞを切り出す
    s_High = (s_Value And &HFFFF0000) \ &H10000
    
End Sub



'---------------------------------------------------------------
'
'   LngToByte()
'
'       機能    ：Long型⇒4つのByte型
'       引数    ：分割するLong型ﾃﾞｰﾀ, 1ﾊﾞｲﾄ目, 2ﾊﾞｲﾄ目, 3ﾊﾞｲﾄ目, 4ﾊﾞｲﾄ目
'       戻り値  ：分割されたLong型ﾃﾞｰﾀ
'       作成    ：2007/03/06 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub LngToByte(s_Value As Long, s_HighHigh As Byte, s_HighLow As Byte, s_LowHigh As Byte, s_LowLow As Byte)

    Dim HighWord As Integer
    Dim LowWord As Integer

    'Integer型に分割
    Call LngToInt(s_Value, HighWord, LowWord)
    
    'Byte型に分割
    Call IntToByte(HighWord, s_HighHigh, s_HighLow)
    Call IntToByte(LowWord, s_LowHigh, s_LowLow)

End Sub


'---------------------------------------------------------------
'
'   Swapi()
'
'       機能    ：Long型⇒4つのByte型
'       引数    ：分割するLong型ﾃﾞｰﾀ, 1ﾊﾞｲﾄ目, 2ﾊﾞｲﾄ目, 3ﾊﾞｲﾄ目, 4ﾊﾞｲﾄ目
'       戻り値  ：分割されたLong型ﾃﾞｰﾀ
'       作成    ：2007/03/06 by 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Function Swapi(s_Value As Integer) As Integer

    Dim HighByte As Byte
    Dim LowByte As Byte

    'Byte型に分割
    Call IntToByte(s_Value, HighByte, LowByte)
    '反転してInteger型に結合
    Swapi = ByteToInt(LowByte, HighByte)

End Function
