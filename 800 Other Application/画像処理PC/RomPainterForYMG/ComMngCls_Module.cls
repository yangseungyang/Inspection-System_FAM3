VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComMng_Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
'*                                                              *
'*  Mcmd_Module                                                 *
'*                                                              *
'*      作成日  ：2007/02/14                                    *
'*      作成者  ：澤原                                          *
'*      ｺﾒﾝﾄ    ：ComCtrlEx.dllを使用するためのｸﾗｽ              *
'*                                                              *
'*                                                              *
'****************************************************************

'Option Explicit

'---------------------------------------------
'外部API関数使用宣言
'---------------------------------------------
'---Sleep関数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'---ｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙ取得用API
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long



'ComCtrlEx使用宣言
'---ｵﾌﾞｼﾞｪｸﾄ変数を用意
Dim ComMng As COMCTRLEXLib.ComMng
'---構造体
Dim k_ConnectPrm As COMCTRLEXLib.ConnectPrm2
Dim k_F3OpenInfo As COMCTRLEXLib.F3OpenInfo2
Dim k_F3OnlineInfo As COMCTRLEXLib.F3OnlineInfo2
Dim k_F3ModTable As COMCTRLEXLib.F3ModTable


'ﾌﾟﾛﾊﾟﾃｨ変数宣言
'---ﾂｰﾙ名称
Private m_AppName As String
'---ﾂｰﾙ名称
Private m_TargetCpuNo As Integer
'---ｽﾛｯﾄ番号
Private m_ModuleSlot As Integer
'---IPｱﾄﾞﾚｽ
Private m_IPAddress As String
'---COMﾎﾟｰﾄ番号
Private m_ComPort As Integer
'---接続方法
Public Enum ConnectMethods
    Serial = 0
    Modem = 1
    ETHER = 2
    FLnet = 3
    USB = 4
End Enum
Private m_ConnectMethod As ConnectMethods
'---接続状態
Public Enum ConnectStates
    ConnectErr = -1
    NotConnect = 0
    Connecting = 1
    Reset = 2
End Enum
Private m_ConnectState As ConnectStates
'---ｺﾏﾝﾄﾞ実行状態
Public Enum CmdStates
    CMD_OK = 0
    CMD_ERR = -1
End Enum
Private m_CmdState As CmdStates
'---ROM種類
Public Enum RomTypes
    WordType = 0
    ByteType = 1
End Enum
Private m_RomType As RomTypes
'---ｴﾝﾃﾞｨｱﾝ
Public Enum EndianModes
    BigEndian = 0
    LittleEndian = 1
End Enum
Private m_EndianMode As EndianModes


'標準変数宣言
Private IOinfoData(16) As String


'標準変数宣言
Private Const MODIFY_ADDRESS As Long = &HE
Private Const MODIFY_DATA As Long = &H8012
Private Const HIGH_WINDOW_ADDRESS As Long = &H10
Private Const LOW_WINDOW_ADDRESS As Long = &H12


'ﾕｰｻﾞ定義のﾃﾞｰﾀ型
Private Type UnsignedInt    '符号なし2ﾊﾞｲﾄ
    Low As Byte
    High As Byte
End Type
Private Type SignedInt      '符号あり2ﾊﾞｲﾄ
    n As Integer
End Type
Private Type UnsignedLong   '符号なし4ﾊﾞｲﾄ
    LowLow As Byte
    LowHigh As Byte
    HighLow As Byte
    HighHigh As Byte
End Type
Private Type SignedLong     '符号付き4ﾊﾞｲﾄ
    n As Long
End Type



'---------------------------------------------------------------
'
'   ﾌﾟﾛﾊﾟﾃｨﾌﾟﾛｼｰｼﾞｬ群
'
'       機能    ：
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/14 By 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
'---ｱﾌﾟﾘｹｰｼｮﾝ名称（設定のみ）
Public Property Let AppName(ByVal newAppName As String)
    m_AppName = newAppName
End Property

'---CPU番号（設定のみ）
Public Property Let TargetCpuNo(ByVal newTargetCpuNo As Integer)
    m_TargetCpuNo = newTargetCpuNo
End Property

'---接続方法
Public Property Get ConnectMethod() As ConnectMethods
    ConnectMethod = m_ConnectMethod
End Property

Public Property Let ConnectMethod(ByVal newConnectMethod As ConnectMethods)
    m_ConnectMethod = newConnectMethod
End Property

'---通信ﾎﾟｰﾄ
Public Property Get ComPort() As Integer
    ComPort = m_ComPort
End Property

Public Property Let ComPort(ByVal newComPort As Integer)
    m_ComPort = newComPort
End Property

'---接続状況
Public Property Get ConnectState() As ConnectStates
    ConnectState = m_ConnectState
End Property

Public Property Let ConnectState(ByVal newConnectState As ConnectStates)
    m_ConnectState = newConnectState
End Property

'---IPｱﾄﾞﾚｽ
Public Property Get IPAddress() As String
    IPAddress = m_IPAddress
End Property

Public Property Let IPAddress(ByVal newIPAddress As String)
    m_IPAddress = newIPAddress
End Property

'---ｽﾛｯﾄ番号
Public Property Get ModuleSlot() As Integer
    ModuleSlot = m_ModuleSlot
End Property

Public Property Let ModuleSlot(ByVal newModuleSlot As Integer)
    m_ModuleSlot = newModuleSlot
End Property

'---ｺﾏﾝﾄﾞ実行状態（設定のみ）
Public Property Get CmdState() As CmdStates
    CmdState = m_CmdState
End Property

'Public Property Let CmdState(ByVal newCmdState As CmdStates)
'    m_CmdState = newCmdState
'End Property

'---ROM種類
Public Property Get RomType() As RomTypes
    RomType = m_RomType
End Property

Public Property Let RomType(ByVal newRomType As RomTypes)
    m_RomType = newRomType
End Property

'---ｴﾝﾃﾞｨｱﾝﾓｰﾄﾞ
Public Property Get EndianMode() As EndianModes
    EndianMode = m_EndianMode
End Property

Public Property Let EndianMode(ByVal newEndianMode As EndianModes)
    m_EndianMode = newEndianMode
End Property



'---------------------------------------------------------------
'
'   ComCtrlLibInit()
'
'       機能    ：通信ﾗｲﾌﾞﾗﾘの使用開始
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/14 By 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub ComCtrlLibInit()

    'ComMngｸﾗｽのｲﾝｽﾀﾝｽ生成
    Set ComMng = New ComMng
    
    '通信ﾗｲﾌﾞﾗﾘ使用準備
    '---ﾌﾟﾛﾊﾟﾃｨ設定(ｱﾌﾟﾘｹｰｼｮﾝ名称とｳｨﾝﾄﾞｳﾊﾝﾄﾞﾙのみ)
    ComMng.AppName = "RP"
    Dim hwnd As Long
    hwnd = FindWindow(vbNullString, m_AppName)
    ComMng.hwnd = hwnd
    
    '変数初期化
    m_ConnectState = NotConnect
    
End Sub



'---------------------------------------------------------------
'
'   ComCtrlLibEnd()
'
'       機能    ：通信ﾗｲﾌﾞﾗﾘの使用終了
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/14 By 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub ComCtrlLibEnd()

    'ComMngｸﾗｽ解放
    Set ComMng = Nothing

End Sub



'---------------------------------------------------------------
'
'   Connect()
'
'       機能    ：FAM3と接続
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/14 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Comopn2」を使用
'
'---------------------------------------------------------------
Public Sub Connect()

    Dim KindOfSP As Integer

    '通信設定
    Select Case m_ConnectMethod
        Case Serial
            '---接続ﾊﾟﾗﾒｰﾀ
            k_ConnectPrm.nConMedia = 0
            k_ConnectPrm.xRsData.nConnect = 0
            k_ConnectPrm.xRsData.nTimeOut = 1
            k_ConnectPrm.xRsData.nRetry = 2
            k_ConnectPrm.xRsData.nComPort = m_ComPort
            '---接続情報
            k_F3OpenInfo.xMode.nMode = &H110
            k_F3OpenInfo.xMode.nDevPara = 0
            k_F3OpenInfo.xMode.nTimeOut = 1
            k_F3OpenInfo.xMode.nRetry = 2
            k_F3OpenInfo.xMode.nPortNo = m_ComPort
            k_F3OpenInfo.xMode.nCpuNo = m_TargetCpuNo
        
        Case ETHER
            '---接続ﾊﾟﾗﾒｰﾀ
            k_ConnectPrm.nConMedia = 2
            k_ConnectPrm.xEther.strTarget = m_IPAddress
            k_ConnectPrm.xEther.nCpuNo = m_TargetCpuNo
            k_ConnectPrm.xEther.nCnTmOut = 10
            '---接続情報
            k_F3OpenInfo.xMode.nMode = &H300
            k_F3OpenInfo.xMode.nIpAddress = Val(m_IPAddress)
            k_F3OpenInfo.xMode.nCpuNo = m_TargetCpuNo
            k_F3OpenInfo.xMode.nConTimeout = 10
       
    End Select


    On Error GoTo Error_Connect

        '他のToolが接続していた時の処理
        Call ComMng.OnlineInf2(k_F3OnlineInfo)
        If k_F3OnlineInfo.xMode.nMode <> 0 Then
'            Ret = MsgBox("他のﾂｰﾙが接続しています。同じ通信条件でのみ接続できますがよろしいですか？", vbQuestion + vbYesNo, "How about it?")
'            If Ret = vbNo Then
'                MsgBox "接続をｷｬﾝｾﾙします。", vbExclamation, "CANCELED"
'                Exit Sub
'            End If
            MsgBox "他のﾂｰﾙが接続しているので接続できません", vbExclamation, "SORRY"
            m_ConnectState = ConnectErr
            Exit Sub
        
        End If
            
        '接続情報ﾊﾟﾗﾒｰﾀ設定
        Call ComMng.SetParam2(k_ConnectPrm)
        
        'FA-M3と接続
        Call ComMng.Comopn2(k_F3OpenInfo)
        m_ConnectState = Connecting
        
        'ｼﾘｱﾙ通信の場合一時的に最速にする
        If m_ConnectMethod = Serial Then
            If m_ConnectState <> Reset Then              'ﾘｾｯﾄの場合は行わない
                Call ComMng.OnlineInf2(k_F3OnlineInfo)
                KindOfSP = Mid(k_F3OnlineInfo.xCpu.strSysID, 5, 2)
                If (KindOfSP = 21) Or (KindOfSP = 25) Or (KindOfSP = 35) Then
                    Call ComMng.RsSpeedChange(2)    '19200bpsﾊﾟﾘﾃｨ偶数
                Else
                    Call ComMng.RsSpeedChange(8)    '115200bpsﾊﾟﾘﾃｨ偶数
                End If
            End If
        End If
        
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_Connect:
    m_ConnectState = ConnectErr
'------------------

End Sub



'---------------------------------------------------------------
'
'   DisConnect()
'
'       機能    ：FAM3から切断
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/14 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Comcls」を使用
'
'---------------------------------------------------------------
Public Sub DisConnect()

    On Error GoTo Error_Comcls
        'ｼﾘｱﾙ通信の場合速度を元に戻す
        If m_ConnectMethod = Serial Then
            If m_ConnectState <> Reset Then     'ﾘｾｯﾄの場合は行わない
                Call ComMng.RsSpeedChange(k_F3OnlineInfo.xMode.nDevPara)
            End If
        End If
        '切断
        ComMng.Comcls
        m_ConnectState = NotConnect

    Exit Sub

'ｴﾗｰ処理-----------
Error_Comcls:
    m_ConnectState = ConnectErr
'------------------

End Sub



'---------------------------------------------------------------
'
'   SystemReset()
'
'       機能    ：CPUﾘｽﾀｰﾄ
'       引数    ：-1:全CPUﾘｽﾀｰﾄ／0:自CPUﾘｽﾀｰﾄ／1～4:指定CPUｼｽﾀｰﾄ
'       戻り値  ：
'       作成    ：2007/05/15 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Modstp」を使用
'
'---------------------------------------------------------------
Public Sub SystemReset(s_CpuNo As Integer)
    
    '停止ﾓｰﾄﾞにする
    Call ComMng.Sysrst(s_CpuNo)
    Sleep (2000)
    
    m_CmdState = CMD_OK

End Sub



'---------------------------------------------------------------
'
'   HaltSP()
'
'       機能    ：SPの運転を停止ﾓｰﾄﾞにする
'       引数    ：
'       戻り値  ：
'       作成    ：2007/02/15 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Modstp」を使用
'
'---------------------------------------------------------------
Public Sub HaltSP()
    
    On Error GoTo Error_HaltSP
        
        '停止ﾓｰﾄﾞにする
        Call ComMng.Modstp
        m_CmdState = CMD_OK
        
    Exit Sub

'ｴﾗｰ処理-----------
Error_HaltSP:
    'MsgBox "HaltSP ERROR", vbCritical, "ERROR"
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   GetIOinfo()
'
'       機能    ：ﾓｼﾞｭｰﾙ構成を取得
'       引数    ：m_IOinfo()
'       戻り値  ：
'       作成    ：2007/02/14 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Cpumod」を使用
'
'---------------------------------------------------------------
Public Sub GetIOinfo(s_IOinfo() As String)

    Dim UnitNo As Integer
    Dim i As Integer
    
    '設定
    UnidNo = 0
    
    On Error GoTo Error_GetIOinfo
        'ﾓｼﾞｭｰﾙ構成を取得
        Call ComMng.Cpumod(UnidNo, k_F3ModTable)
        'ﾓｼﾞｭｰﾙIDを返す
        For i = 0 To 15
            s_IOinfo(i) = k_F3ModTable.xMod(i).strModID
        Next i
        '戻り値
        m_CmdState = CMD_OK

    Exit Sub

'ｴﾗｰ処理-----------
Error_GetIOinfo:
    'MsgBox "GetIOinfo ERROR", vbCritical, "ERROR"
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   WriteGAReg()
'
'       機能    ：GAの内部ﾚｼﾞｽﾀの$200番地以降へﾗｲﾄ(ﾜｰﾄﾞ単位)
'       引数    ：相対ｱﾄﾞﾚｽ, ﾗｲﾄﾃﾞｰﾀ(word)
'       戻り値  ：
'       作成    ：2007/02/16 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「MemwrtByte/Word」を使用
'
'                 GAの内部ﾚｼﾞｽﾀ($200～)は$7C0000から
'                 ｽﾛｯﾄ毎に割り当てられている。
'
'                 GateArraywrtﾒｿｯﾄﾞを使用したかったが、
'                 RDYのついてないﾓｼﾞｭｰﾙに対しても行うため、
'                 ｲﾆｼｬﾗｲｽﾞ処理を行わなければならず、諦め・・。
'
'---------------------------------------------------------------
Public Sub WriteGAReg(s_RegisterNo As Long, s_WriteData As Long)

    Dim SlotAddress As Long, AccessAddress As Long
    Dim HighWord As Integer, LowWord As Integer

    SlotAddress = &H7C0000 + (m_ModuleSlot - 1) * 128
    
    On Error GoTo Error_WriteGAReg
        'ｲﾆｼｬﾗｲｽﾞ処理(RDYのついていないﾓｼﾞｭｰﾙのために必要)
        '---GAｲﾆｼｬﾗｲｽﾞﾚｼﾞｽﾀ($21B)に$80をﾗｲﾄ
        AccessAddress = SlotAddress + &H1A
        Call ComMng.MemwrtWord(1, AccessAddress, &H8080)
        '---GAｽﾚｰﾌﾞｽﾃｰﾀｽﾚｼﾞｽﾀ($278)に$00をﾗｲﾄ
        AccessAddress = SlotAddress + &H78
        Call ComMng.MemwrtWord(1, AccessAddress, &H0)
        '---指定のGAﾚｼﾞｽﾀに値をﾗｲﾄ
        If s_RegisterNo < 0 Then
            s_RegisterNo = s_RegisterNo + 65536
        End If
        AccessAddress = SlotAddress + s_RegisterNo
        If s_WriteData < 0 Then
            s_WriteData = s_WriteData + 65536
        End If
        Call LngToInt(s_WriteData, HighWord, LowWord)   '下位ﾜｰﾄﾞを取り出す
        Call ComMng.MemwrtWord(1, AccessAddress, LowWord)
        m_CmdState = CMD_OK
    
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_WriteGAReg:
    'MsgBox ("WriteGAReg ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   WriteGAInternalReg()
'
'       機能    ：GAの内部ﾚｼﾞｽﾀの$001～1FE番地へﾗｲﾄ(ﾜｰﾄﾞ単位)
'       引数    ：相対ｱﾄﾞﾚｽ, ﾗｲﾄﾃﾞｰﾀ(word)
'       戻り値  ：
'       作成    ：2007/02/16 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「MemwrtWord」を使用
'
'                 GAの内部ﾚｼﾞｽﾀ($001～1FE)は割当てられていない。
'                 よって、$720000～64KBのCPU2,3の空間にﾏｯﾋﾟﾝｸﾞする。
'
'---------------------------------------------------------------
Public Sub WriteGAInternalReg(s_RegisterNo As Long, s_WriteData As Long)
    
    Dim AccessAddress As Long
    Dim HighWord As Integer, LowWord As Integer

    'ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽﾚｼﾞｽﾀ設定(ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽ:$20000、64KB)
    Call WriteGAReg(MODIFY_ADDRESS, MODIFY_DATA)
    If (m_CmdState = CMD_ERR) Then Exit Sub
    
    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定(KP内部ﾚｼﾞｽﾀｱｸｾｽ)
    If m_RomType = WordType Then
        Call WriteGAReg(LOW_WINDOW_ADDRESS, &H18)
    Else
        Call WriteGAReg(LOW_WINDOW_ADDRESS, &H12)
    End If
    If (m_CmdState = CMD_ERR) Then Exit Sub
    
    '$10000毎に切替え
    If s_RegisterNo < 0 Then
        s_RegisterNo = s_RegisterNo + 65536
    End If
    AccessAddress = &H760000 + s_RegisterNo
    
    '下位ﾜｰﾄﾞを取り出す
    If s_WriteData < 0 Then
        s_WriteData = s_WriteData + 65536
    End If
    Call LngToInt(s_WriteData, HighWord, LowWord)
    
    '指定のGA内部ﾚｼﾞｽﾀに値をﾗｲﾄ
    On Error GoTo Error_WriteGAInternalReg
        Call ComMng.MemwrtWord(1, AccessAddress, LowWord)
        m_CmdState = CMD_OK
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_WriteGAInternalReg:
    'MsgBox ("WriteGAInternalReg ERROR")
    m_CmdState = CMD_ERR
'------------------
    
End Sub



'---------------------------------------------------------------
'
'   WriteSpInternalMemory()
'
'       機能    ：SPの内部ﾒﾓﾘに1ﾊﾞｲﾄ書込み
'       引数    ：ｱﾄﾞﾚｽ, ﾗｲﾄﾃﾞｰﾀ(Byte型)
'       戻り値  ：
'       作成    ：2007/06/22 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「MemwrtByte」を使用
'
'---------------------------------------------------------------
Public Sub WriteSpInternalMemory(s_Address As Long, s_WriteData As Byte)

    '指定のｱﾄﾞﾚｽに値をﾗｲﾄ
    On Error GoTo Error_WriteSpInternalMemory
        Call ComMng.MemwrtByte(1, s_Address, s_WriteData)
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_WriteSpInternalMemory:
    'MsgBox ("WriteSpInternalMemory ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub




'---------------------------------------------------------------
'
'   WriteMemoryByte()
'
'       機能    ：ﾓｼﾞｭｰﾙ内ﾒﾓﾘに1ﾜｰﾄﾞ書込み
'       引数    ：ｱﾄﾞﾚｽ, ﾗｲﾄﾃﾞｰﾀ(Byte型)
'       戻り値  ：
'       作成    ：2007/04/20 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「MemwrtByte」を使用
'
'                 $10000番地毎にｳｨﾝﾄﾞｳｱﾄﾞﾚｽを切り替える。
'
'---------------------------------------------------------------
Public Sub WriteMemoryByte(s_Address As Long, s_WriteData As Byte)

    Dim HighWindowData As Long, LowWindowData As Long
    Dim AccessAddress As Long

    'ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽﾚｼﾞｽﾀ設定(ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽ:$20000、64KB)
    Call WriteGAReg(MODIFY_ADDRESS, MODIFY_DATA)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定
    HighWindowData = (s_Address \ &H40000) * 4
    Call WriteGAReg(HIGH_WINDOW_ADDRESS, HighWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub
    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定(ﾓｼﾞｭｰﾙ内部ﾃﾞﾊﾞｲｽｱｸｾｽ)
    If m_RomType = WordType Then
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H10
    Else
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H12
    End If
    Call WriteGAReg(LOW_WINDOW_ADDRESS, LowWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '$10000毎に切替え
    If s_Address < 0 Then
        s_Address = s_Address + 65536
    End If
    AccessAddress = &H760000 + (s_Address Mod &H10000)
        
    '指定のｱﾄﾞﾚｽに値をﾗｲﾄ
    On Error GoTo Error_WriteMemoryByte
        Call ComMng.MemwrtByte(1, AccessAddress, s_WriteData)
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_WriteMemoryByte:
    'MsgBox ("WriteMemoryByte ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   WriteMemoryWord()
'
'       機能    ：ﾓｼﾞｭｰﾙ内ﾒﾓﾘに1ﾜｰﾄﾞ書込み
'       引数    ：ｱﾄﾞﾚｽ, ﾗｲﾄﾃﾞｰﾀ(Integer型)
'       戻り値  ：
'       作成    ：2007/04/20 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「MemwrtWord」を使用
'
'                 $10000番地毎にｳｨﾝﾄﾞｳｱﾄﾞﾚｽを切り替える。
'
'---------------------------------------------------------------
Public Sub WriteMemoryWord(s_Address As Long, s_WriteData As Integer)

    Dim HighWindowData As Long, LowWindowData As Long
    Dim AccessAddress As Long

    'ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽﾚｼﾞｽﾀ設定(ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽ:$20000、64KB)
    Call WriteGAReg(MODIFY_ADDRESS, MODIFY_DATA)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定
    HighWindowData = (s_Address \ &H40000) * 4
    Call WriteGAReg(HIGH_WINDOW_ADDRESS, HighWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定(ﾓｼﾞｭｰﾙ内部ﾃﾞﾊﾞｲｽｱｸｾｽ)
    If m_RomType = WordType Then
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H10
    Else
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H12
    End If
    Call WriteGAReg(LOW_WINDOW_ADDRESS, LowWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '$10000毎に切替え
    If s_Address < 0 Then
        s_Address = s_Address + 65536
    End If
    AccessAddress = &H760000 + (s_Address Mod &H10000)  '$10000毎に切替え
        
    '指定のｱﾄﾞﾚｽに値をﾗｲﾄ
    On Error GoTo Error_WriteMemoryWord
        Call ComMng.MemwrtWord(1, AccessAddress, s_WriteData)
        m_CmdState = CMD_OK
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_WriteMemoryWord:
    'MsgBox ("WriteMemoryWord ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   ReadMemoryByte()
'
'       機能    ：ﾓｼﾞｭｰﾙ内ﾒﾓﾘを1ﾊﾞｲﾄ読出し
'       引数    ：ｱﾄﾞﾚｽ, 読出し変数(byte型)
'       戻り値  ：
'       作成    ：2007/04/20 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「MemredByte」を使用
'
'                 $10000番地毎にｳｨﾝﾄﾞｳｱﾄﾞﾚｽを切り替える。
'
'---------------------------------------------------------------
Public Sub ReadMemoryByte(s_Address As Long, s_ReadData As Byte)

    Dim HighWindowData As Long, LowWindowData As Long
    Dim AccessAddress As Long

    'ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽﾚｼﾞｽﾀ設定(ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽ:$20000、64KB)
    Call WriteGAReg(MODIFY_ADDRESS, MODIFY_DATA)
    If (m_CmdState = CMD_ERR) Then Exit Sub
    
    '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定
    HighWindowData = (s_Address \ &H40000) * 4
    Call WriteGAReg(HIGH_WINDOW_ADDRESS, HighWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定(ﾓｼﾞｭｰﾙ内部ﾃﾞﾊﾞｲｽｱｸｾｽ)
    If m_RomType = WordType Then
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H10
    Else
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H12
    End If
    Call WriteGAReg(LOW_WINDOW_ADDRESS, LowWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '指定のｱﾄﾞﾚｽに値をﾗｲﾄ
    On Error GoTo Error_ReadMemoryByte
    
        If s_Address < 0 Then
            s_Address = s_Address + 65536
        End If
        AccessAddress = &H760000 + (s_Address Mod &H10000)  '$10000毎に切替え
        
        Call ComMng.MemredByte(1, AccessAddress, s_ReadData)
        m_CmdState = CMD_OK
    
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_ReadMemoryByte:
    'MsgBox ("ReadMemoryByte ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   ReadMemoryWord()
'
'       機能    ：ﾓｼﾞｭｰﾙ内ﾒﾓﾘを読出し
'       引数    ：ｱﾄﾞﾚｽ, 読出し変数(Integer型)
'       戻り値  ：
'       作成    ：2007/04/20 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「MemredWord」を使用
'
'                 $10000番地毎にｳｨﾝﾄﾞｳｱﾄﾞﾚｽを切り替える。
'
'---------------------------------------------------------------
Public Sub ReadMemoryWord(s_Address As Long, s_ReadData As Integer)

    Dim HighWindowData As Long, LowWindowData As Long
    Dim AccessAddress As Long

    'ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽﾚｼﾞｽﾀ設定(ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽ:$20000、64KB)
    Call WriteGAReg(MODIFY_ADDRESS, MODIFY_DATA)
    If (m_CmdState = CMD_ERR) Then Exit Sub
    
    '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定
    HighWindowData = (s_Address \ &H40000) * 4
    Call WriteGAReg(HIGH_WINDOW_ADDRESS, HighWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定(ﾓｼﾞｭｰﾙ内部ﾃﾞﾊﾞｲｽｱｸｾｽ)
    If m_RomType = WordType Then
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H10
    Else
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H12
    End If
    Call WriteGAReg(LOW_WINDOW_ADDRESS, LowWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '指定のｱﾄﾞﾚｽに値をﾗｲﾄ
    On Error GoTo Error_ReadMemoryWord
    
        If s_Address < 0 Then
            s_Address = s_Address + 65536
        End If
        AccessAddress = &H760000 + (s_Address Mod &H10000)  '$10000毎に切替え
        
        Call ComMng.MemredWord(1, AccessAddress, s_ReadData)
        m_CmdState = CMD_OK
    
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_ReadMemoryWord:
    'MsgBox ("ReadMemoryWord ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   CancelMap()
'
'       機能    ：ﾓｼﾞｭｰﾙのﾏｯﾋﾟﾝｸﾞを解除する
'       引数    ：
'       戻り値  ：
'       作成    ：2007/03/13 By 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub CancelMap()

    Dim SlotAddress As Long, AccessAddress As Long

    SlotAddress = &H7C0000 + (m_ModuleSlot - 1) * 128
    AccessAddress = SlotAddress + MODIFY_ADDRESS

    '指定のｱﾄﾞﾚｽに値をﾗｲﾄ
    On Error GoTo Error_CancelMap
        Call ComMng.MemwrtWord(1, AccessAddress, 0)
        m_CmdState = CMD_OK
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_CancelMap:
    'MsgBox ("CancelMap ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub
    


'---------------------------------------------------------------
'
'   RdyLedOff()
'
'       機能    ：ﾓｼﾞｭｰﾙのRDY LEDを消灯する。
'       引数    ：
'       戻り値  ：
'       作成    ：2007/03/13 By 澤原
'       ｺﾒﾝﾄ    ：
'
'---------------------------------------------------------------
Public Sub RdyLedOff()

    Dim SlotAddress As Long
    Dim AccessAddress As Long

    SlotAddress = &H7C0000 + (m_ModuleSlot - 1) * 128
    
    On Error GoTo Error_RdyLedOff
    
        'GAｲﾆｼｬﾗｲｽﾞﾚｼﾞｽﾀ($21B)に$0をﾗｲﾄ
        AccessAddress = SlotAddress + &H1A
        Call ComMng.MemwrtWord(1, AccessAddress, &H0)
    
    Exit Sub

'ｴﾗｰ処理-----------
Error_RdyLedOff:
    MsgBox ("RdyLedOff ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub
    
    

'---------------------------------------------------------------
'
'   SectorEraseRom()
'
'       機能    ：ﾓｼﾞｭｰﾙのROMを消去
'       引数    ：消去先頭ｱﾄﾞﾚｽ
'       戻り値  ：
'       作成    ：2007/02/23 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Flashdel」を使用
'
'---------------------------------------------------------------
Public Sub SectorEraseRom(s_Address As Long)

    If s_Address < 0 Then
        s_Address = s_Address + 65536
    End If
    
    '0でFlashdelをｺｰﾙすると全消去になるので、+2に投げる
    If s_Address = 0 Then
        s_Address = s_Address + 2
    End If
   
   'ROM消去ﾒｿｯﾄﾞ
    On Error GoTo Error_SectorEraseRom
        Call ComMng.Flashdel(m_RomType, 0, m_ModuleSlot, s_Address)
        m_CmdState = CMD_OK
    
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_SectorEraseRom:
    'MsgBox ("SectorEraseRom Error!")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   TipEraseRom()
'
'       機能    ：ﾓｼﾞｭｰﾙのROMを消去
'       引数    ：FlashROMの先頭ｱﾄﾞﾚｽ
'       戻り値  ：
'       作成    ：2007/05/28 By 澤原
'       ｺﾒﾝﾄ    ：Flash系のMｺﾏﾝﾄﾞを使用せずに、
'               　直接Flashにｱｸｾｽしてﾁｯﾌﾟ消去ｺﾏﾝﾄﾞ実行
'               　Flashdelのﾀｲﾑｱｳﾄは8秒なので、
'               　8秒で終わらないもののためにFlashdelを使用しない
'
'---------------------------------------------------------------
Public Sub TipEraseRom(s_Address As Long, s_TimeoutSec As Integer)

    Dim HighWindowData As Long, LowWindowData As Long
    Dim i_ReadData As Integer, CheckCounter As Integer

    'ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽﾚｼﾞｽﾀ設定(ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽ:$20000、64KB)
    Call WriteGAReg(MODIFY_ADDRESS, MODIFY_DATA)
    If (m_CmdState = CMD_ERR) Then Exit Sub
    
    '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定
    HighWindowData = (s_Address \ &H40000) * 4
    Call WriteGAReg(HIGH_WINDOW_ADDRESS, HighWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定(ﾓｼﾞｭｰﾙ内部ﾃﾞﾊﾞｲｽｱｸｾｽ)
    If m_RomType = WordType Then
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H10
    Else
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H12
    End If
    Call WriteGAReg(LOW_WINDOW_ADDRESS, LowWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub
    
   'ROMﾁｯﾌﾟ消去ｼｰｹﾝｽ
    On Error GoTo Error_TipEraseRom
        If m_RomType = WordType Then
            Call ComMng.MemwrtWord(1, &H76AAAA, &HAAAA)
            Call ComMng.MemwrtWord(1, &H765554, &H5555)
            Call ComMng.MemwrtWord(1, &H76AAAA, &H8080)
            Call ComMng.MemwrtWord(1, &H76AAAA, &HAAAA)
            Call ComMng.MemwrtWord(1, &H765554, &H5555)
            Call ComMng.MemwrtWord(1, &H76AAAA, &H1010)
        Else
            Call ComMng.MemwrtByte(1, &H76AAAA, &HAA)
            Call ComMng.MemwrtByte(1, &H765555, &H55)
            Call ComMng.MemwrtByte(1, &H76AAAA, &H80)
            Call ComMng.MemwrtByte(1, &H76AAAA, &HAA)
            Call ComMng.MemwrtByte(1, &H765555, &H55)
            Call ComMng.MemwrtByte(1, &H76AAAA, &H10)
        End If
    
    'ROM全消去完了確認
    s_TimeoutSec = s_TimeoutSec * 3
    CheckCounter = 0
    Do While i_ReadData <> &HFFFF
        DoEvents
        Call ReadMemoryWord(s_Address, i_ReadData)
        'Debug.Print Hex(i_ReadData)
        If EraseMode = Error_Erase Then Exit Do
        Sleep (200)
        CheckCounter = CheckCounter + 1
        If CheckCounter > s_TimeoutSec Then  'ﾀｲﾑｱｳﾄ
            m_CmdState = CMD_ERR
            Exit Do
        End If
        'ﾓｼﾞｭｰﾙのないｽﾛｯﾄに投げた時はｴﾗｰ
'        If i_ReadData = &H5345 Then
'            m_CmdState = CMD_ERR
'            Exit Do
'        End If
    Loop
    
    Exit Sub
    
   
'ｴﾗｰ処理-----------
Error_TipEraseRom:
    'MsgBox ("TipEraseRom Error!")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   WriteRom()
'
'       機能    ：ﾓｼﾞｭｰﾙのROMに書込み
'       引数    ：書込み先頭ｱﾄﾞﾚｽ, 書込みｻｲｽﾞ, 書込みﾃﾞｰﾀ(Byte型配列)
'       戻り値  ：
'       作成    ：2007/02/23 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Flashwrt」を使用
'
'---------------------------------------------------------------
Public Sub WriteRom(s_Address As Long, s_WriteSize As Integer, s_WriteData() As Byte)

    Dim DataBuff As Byte
    
    If s_Address < 0 Then
        s_Address = s_Address + 65536
    End If
        
    'ｴﾝﾃﾞｨｱﾝ変換
    If m_EndianMode = LittleEndian Then
        For i = 0 To (s_WriteSize / 2 - 1)
            DataBuff = s_WriteData(i * 2)
            s_WriteData(i * 2) = s_WriteData(i * 2 + 1)
            s_WriteData(i * 2 + 1) = DataBuff
        Next
    End If
        
    'ROM書込みﾒｿｯﾄﾞ
    On Error GoTo Error_WriteRom
        Call ComMng.Flashwrt(m_RomType, 0, m_ModuleSlot, s_Address, s_WriteSize, s_WriteData(0))
        m_CmdState = CMD_OK
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_WriteRom:
    'MsgBox ("WriteRom ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   ReadRom()
'
'       機能    ：ﾓｼﾞｭｰﾙのROMを読出し
'       引数    ：読出し先頭ｱﾄﾞﾚｽ, 読出しｻｲｽﾞ, 読出しﾃﾞｰﾀ(Byte型配列)
'       戻り値  ：
'       作成    ：2007/04/20 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Flashwrt」を使用
'
'---------------------------------------------------------------
Public Sub ReadRom(s_Address As Long, s_DataNum As Integer, s_ReadData() As Byte)

    Dim HighWindowData As Long, LowWindowData As Long
    Dim AccessAddress As Long
    Dim i As Integer
    Dim DataBuff As Byte

    'ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽﾚｼﾞｽﾀ設定(ﾓﾃﾞｨﾌｧｲｱﾄﾞﾚｽ:$20000、64KB)
    Call WriteGAReg(MODIFY_ADDRESS, MODIFY_DATA)
    If (m_CmdState = CMD_ERR) Then Exit Sub
    
    '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定
    HighWindowData = (s_Address \ &H40000) * 4
    Call WriteGAReg(HIGH_WINDOW_ADDRESS, HighWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定(ﾓｼﾞｭｰﾙ内部ﾃﾞﾊﾞｲｽｱｸｾｽ)
    If m_RomType = WordType Then
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H10
    Else
        LowWindowData = ((s_Address Mod &H40000) \ &H10000) * 16384 + &H12
    End If
    Call WriteGAReg(LOW_WINDOW_ADDRESS, LowWindowData)
    If (m_CmdState = CMD_ERR) Then Exit Sub

    '$10000毎に切替え
    If s_Address < 0 Then
        s_Address = s_Address + 65536
    End If
    AccessAddress = &H760000 + (s_Address Mod &H10000)
  
    'MAXは512ﾊﾞｲﾄ
    If s_DataNum > 512 Then
        s_DataNum = 512
    End If

    '指定のｱﾄﾞﾚｽに値をﾗｲﾄ
    On Error GoTo Error_ReadRom
        Call ComMng.MemredByte(s_DataNum, AccessAddress, s_ReadData(0))
        m_CmdState = CMD_OK
        'ｴﾝﾃﾞｨｱﾝ変換
        If m_EndianMode = LittleEndian Then
            For i = 0 To (s_DataNum / 2 - 1)
                DataBuff = s_ReadData(i * 2)
                s_ReadData(i * 2) = s_ReadData(i * 2 + 1)
                s_ReadData(i * 2 + 1) = DataBuff
            Next
        End If
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_ReadRom:
    'MsgBox ("ReadRom ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub



'---------------------------------------------------------------
'
'   CopyRom()
'
'       機能    ：ﾓｼﾞｭｰﾙのROMをｺﾋﾟｰ
'       引数    ：ｺﾋﾟｰ元ｽﾛｯﾄ,ｺﾋﾟｰ先ｽﾛｯﾄ,ｺﾋﾟｰ先頭ｱﾄﾞﾚｽ, ｺﾋﾟｰｻｲｽﾞ)
'       戻り値  ：
'       作成    ：2007/05/07 By 澤原
'       ｺﾒﾝﾄ    ：ComCtrlExのI/F関数「Flashcopy」を使用
'
'---------------------------------------------------------------
Public Sub CopyRom(srcSlot As Integer, tarSlot As Integer, s_Address As Long, s_CopySize As Integer)

    If s_Address < 0 Then
        s_Address = s_Address + 65536
    End If
        
    'ROMｺﾋﾟｰﾒｿｯﾄﾞ
    On Error GoTo Error_CopyRom
        Call ComMng.Flashcopy(m_RomType, s_CopySize, srcSlot, tarSlot, s_Address)
        m_CmdState = CMD_OK
    Exit Sub
    
'ｴﾗｰ処理-----------
Error_CopyRom:
    'MsgBox ("CopyRom ERROR")
    m_CmdState = CMD_ERR
'------------------

End Sub

