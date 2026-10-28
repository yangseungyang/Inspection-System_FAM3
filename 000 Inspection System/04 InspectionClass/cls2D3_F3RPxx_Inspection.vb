Option Strict On : Option Explicit On

'RPxxのIOﾊﾞｽ検査とﾁｪｯｸｻﾑ照合、主銘板印字

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm020_InspectionStartEnd
Imports System.Text
Imports System.Threading


Public Class cls2D3_F3RPxx_Inspection
    'Class ================================================
    Dim cSP As cls121_SP_TEST_MODE
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cLnk As cls080_F3LinkToLE
    Dim cWin As cls001_Window_Control
    Dim cMac As cls031_MacAddress
    Dim cSer As cls030_LineSerial
    Dim cLbl As cls011_Excel_Operation
    Dim cMcm As cls120_SP_M_Command
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl
    Dim dfrm As frm500_MessagePopUp
    Dim cTEST As cls121_SP_TEST_MODE
    Dim cKP As cls124_KP_IOBUS_TEST

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String

    Dim mLOG As String

    Dim IP_U2 As String = "192.168.1.11"        'B側ﾒｲﾝCPU

    'Property =====================================================================================
    'Error message --------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public ReadOnly Property ERR_CODE() As String
        Get
            ERR_CODE = mErrorCode
        End Get
    End Property

    Public Sub TestExecute(ByVal iName As String)

        Select Case Mid(iName, 1, 3)

            Case "I01" : Call I01(iName)            'I01:MACｱﾄﾞﾚｽ読込みｽﾛｯﾄ2（RP4x5x）
            Case "I02" : Call I02(iName)            'I02:MODEL読込み照合ｽﾛｯﾄ2（RP4x5x）
            Case "I03" : Call I03(iName)            'I03:ﾁｪｯｸｻﾑ読込み照合ｽﾛｯﾄ2（RP4x5x）

            Case "I50" : Call I50(iName)            'I50:SLOT/XDFAIL TEST
            Case "I51" : Call I51(iName)            'I51:ACS SIGNAL TESTｽﾛｯﾄ5
            Case "I52" : Call I52(iName)            'I52:BCLR TESTｽﾛｯﾄ5
            Case "I53" : Call I53(iName)            'I53:BR TESTｽﾛｯﾄ5
            Case "I55" : Call I55(iName)            'I55:I/O BUS INTERRUPT TESTｽﾛｯﾄ5
            Case "I56" : Call I56(iName)            'I56:XLFAIL TESTｽﾛｯﾄ5
            Case "I57" : Call I57(iName)            'I57:XPFAIL TESTｽﾛｯﾄ5
            Case "I58" : Call I58(iName)            'I58:XSFAIL TESTｽﾛｯﾄ5
            Case "I59" : Call I59(iName)            'I59:IOﾊﾞｽ検査前準備

            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF
            Case "I92" : Call I92(iName)            'I92:検査開始時処理
            Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save        'Rev 0.26 KSJ ADD
            Case "IP1" : Call IP1(iName)            'IP1:PrintLabel
            Case "IP2" : Call IP2(iName)            'IP2:OK PrintLabel

            Case "IA0"                              'IA0:ｼﾘｱﾙ番号発番確認
                'If SerialRead() = "" Then
                'End If

                'Case Else : mErrorComment = "検査プロシージャが定義されていません。"   'Rev0.11DEL
            Case Else : mErrorComment = "검사 프로시져가 정의 되어있지 않습니다."   'Rev0.11ADD
                MsgBox(mErrorComment)

        End Select

    End Sub

    Public Sub I01(ByVal iName As String)
        '==============================================================================================================
        ' Function : I01:MACｱﾄﾞﾚｽ読込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Dim Model As String = ModelRead()
        Dim mMAC1 As String
        Dim mMAC2 As String
        Dim mMAC3 As String
        Dim mMAC As String

        Call TestPreparation(0) : Call WriteLog(False, mLOG) : Call mSecWait(300)

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        If InStr(Model, "F3RP55") <> 0 Then Call WaitStart(5)

        Try
            Call WriteLog(False, "0x2FF00ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002FF00") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mMAC1 = Strings.Right(cMcm.RecvText, 4)
            'Call mSecWait(100)
            Call WriteLog(False, "0x2FF02ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002FF02") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mMAC2 = Strings.Right(cMcm.RecvText, 4)
            'Call mSecWait(100)
            Call WriteLog(False, "0x2FF04ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002FF04") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mMAC3 = Strings.Right(cMcm.RecvText, 4)
            'Call mSecWait(100)
            If InStr(Model, "-1P") <> 0 Or InStr(Model, "-5P") <> 0 Then
                mMAC = Mid(mMAC1, 3, 2) & Mid(mMAC1, 1, 2) & Mid(mMAC2, 3, 2) & Mid(mMAC2, 1, 2) & Mid(mMAC3, 3, 2) & Mid(mMAC3, 1, 2)
            Else
                mMAC = mMAC1 & mMAC2 & mMAC3
            End If
            If mMAC.Length <> 12 Or Strings.Left(mMAC, 6) <> "000064" Then Throw New ApplicationException(mErrorComment)
            Call WriteLog(False, "MACｱﾄﾞﾚｽ:" & mMAC)
            Call WriteData("MacAddress", mMAC)
        Catch ex As Exception
            mErrorComment = "MACｱﾄﾞﾚｽ読込みｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

    End Sub

    Public Sub I02(ByVal iName As String)
        '==============================================================================================================
        ' Function : I02:MODEL読込み照合
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Dim Model As String = ModelRead()
        Dim mModel1 As String
        Dim mModel2 As String
        Dim mModel3 As String
        Dim mModel As String = ""

        Call TestPreparation(0) : Call WriteLog(False, mLOG)

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        Try
            Call WriteLog(False, "0x2DFF0ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 4, "0002DFF0") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mModel1 = Strings.Right(cMcm.RecvText, 8)
            Call WriteLog(False, "0x2DFF4ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 4, "0002DFF4") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mModel2 = Strings.Right(cMcm.RecvText, 8)
            Call WriteLog(False, "0x2DFF8ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 4, "0002DFF8") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mModel3 = Mid(Strings.Right(cMcm.RecvText, 8), 1, 2)

            For i = 1 To 17 Step 2
                mModel &= Chr(CInt("&H" & Mid(mModel1 & mModel2 & mModel3, i, 2)))
            Next i

            If mModel.Length <> 9 Then Throw New ApplicationException(mErrorComment)
            Call WriteLog(False, "MODEL:" & mModel)
            If Strings.Left(Model, 9) <> mModel Then Throw New ApplicationException(mErrorComment)

        Catch ex As Exception
            mErrorComment = "MODEL読込み照合ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

    End Sub

    Public Sub I03(ByVal iName As String)
        '==============================================================================================================
        ' Function : I03:ﾁｪｯｸｻﾑ読込み照合
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Dim Model As String = ModelRead()
        Dim mReadCheckSum As Integer
        Dim mBASECalcCheckSum As Integer
        Dim mOSBTCalcCheckSum As Integer
        Dim mModel As String = ""

        Dim mFirmCB As String = ""
        Dim mFirmOB As String = ""
        Dim mFirmOS As String = ""
        Dim mFirmCBCheckSum As String = ""
        Dim mFirmOBCheckSum As String = ""
        Dim mFirmOSCheckSum As String = ""

        Dim dbName As String
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim mFirmName() As String = {""}

        '各ﾌｧｰﾑ部番の読込み
        dbName = mAppPath & "520 Product Firmware\RPxx\FirmName.mdb"
        tbName = "tblFirmName"
        rField = "FirmCB,FirmOB,FirmOS"
        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            If InStr(Model, "-5P") <> 0 Then
                sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE MSCODE = '" & Strings.Left(Model, 9) & "'"
            Else
                sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE MSCODE = '" & Strings.Left(Model, 5) & "x-" & Mid(Split(Model, "-")(1), 1, 2) & "'"
            End If
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)
            mFirmCB = sData(0) : mFirmOB = sData(1) : mFirmOS = sData(2)
            Call WriteLog(False, "基本ﾌﾞｰﾄ:" & mFirmCB & ",OSﾌﾞｰﾄ:" & mFirmOB & ",OS:" & mFirmOS)
            mFirmName = sData
        Catch ex As Exception
            mErrorComment = "各ﾌｧｰﾑ部番読込み異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        'ﾁｪｯｸｻﾑの読込み
        Try
            For i = 0 To UBound(mFirmName)
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                tbName = "tblFirmInfo"
                cADO.TB_NAME = tbName

                Select Case mFirmName(i)
                    Case ""
                    Case mFirmOS
                        sSQL = "SELECT Rev FROM " & tbName & " WHERE Model = '" & mFirmName(i) & "'"
                        cADO.DB_SQL = sSQL
                        Call cADO.DB_READ_JET(fName, sData)
                        mFirmOSCheckSum = Strings.Left(mFirmOS, 1) & Strings.Right(mFirmOS, 5) & "0" & sData(0)
                    Case Else
                        sSQL = "SELECT Sum FROM " & tbName & " WHERE Model = '" & mFirmName(i) & "'"
                        cADO.DB_SQL = sSQL
                        Call cADO.DB_READ_JET(fName, sData)
                        If mFirmName(i) = mFirmCB Then mFirmCBCheckSum = sData(0)
                        If mFirmName(i) = mFirmOB Then mFirmOBCheckSum = sData(0)
                End Select
                cADO = Nothing
            Next
            Call WriteLog(False, "基本ﾌﾞｰﾄﾁｪｯｸｻﾑ:" & mFirmCBCheckSum & ",OSﾌﾞｰﾄﾁｪｯｸｻﾑ:" & mFirmOBCheckSum & ",OSﾁｪｯｸ:" & mFirmOSCheckSum)
        Catch ex As Exception
            mErrorComment = "ﾁｪｯｸｻﾑ読込み異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        '基本ﾌﾞｰﾄ
        Call WriteLog(False, "基本ﾌﾞｰﾄ読込み")
        Call TestPreparation(0) : Call WriteLog(False, mLOG)                 '前準備。RPのｱﾄﾞﾚｽ0x00000を0x20000に見えるようにする

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        Try
            '読込み
            Call WriteLog(False, "0x2BFFEﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 4, "0002BFFE") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 4))
            If InStr(Model, "-1P") <> 0 Or InStr(Model, "-5P") <> 0 Or InStr(Model, "-9Z") <> 0 Then
                mBASECalcCheckSum = ((mReadCheckSum >> 8) + (mReadCheckSum << 8) And &HFFFF) + (mReadCheckSum >> 8) + (mReadCheckSum And &HFF)
            Else
                mBASECalcCheckSum = (mReadCheckSum >> 8) + (mReadCheckSum And &HFF) + mReadCheckSum
            End If
            WriteLog(False, "基本ﾌﾞｰﾄCheckSUM:" & Convert.ToString(mBASECalcCheckSum, 16))
            '照合
            If mBASECalcCheckSum <> CInt("&H" & mFirmCBCheckSum) Then
                mErrorComment = "基本ﾌﾞｰﾄCheckSUM照合異常。読込み:" & Convert.ToString(mBASECalcCheckSum, 16) & " FHSUM：" & mFirmCBCheckSum
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "基本ﾌﾞｰﾄ照合OK")
        Catch ex As Exception
            cMcm = Nothing
            If mErrorComment = "" Then mErrorComment = "基本ﾌﾞｰﾄ読込み照合ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        If InStr(Model, "-0P") <> 0 Or InStr(Model, "-1P") <> 0 Or InStr(Model, "-3P") <> 0 Then Exit Sub
        If InStr(Model, "-5P") <> 0 Or InStr(Model, "-9Z") <> 0 Then GoTo OS1

        'OSﾌﾞｰﾄ
        Call WriteLog(False, "OSﾌﾞｰﾄ読込み")
        Call TestPreparation(7) : Call WriteLog(False, mLOG)                '前準備。RPのｱﾄﾞﾚｽ0x70000を0x20000に見えるようにする
        Try
            '読込み
            Call WriteLog(False, "0x2EFFEﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 4, "0002EFFE") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 4))
            mOSBTCalcCheckSum = (mReadCheckSum >> 8) + (mReadCheckSum And &HFF) + mReadCheckSum
            WriteLog(False, "OSﾌﾞｰﾄCheckSUM:" & Convert.ToString(mOSBTCalcCheckSum, 16))
            '照合
            If mOSBTCalcCheckSum <> CInt("&H" & mFirmOBCheckSum) Then
                mErrorComment = "OSﾌﾞｰﾄCheckSUM照合異常。読込み:" & Convert.ToString(mOSBTCalcCheckSum, 16) & " FHSUM：" & mFirmOBCheckSum
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "OSﾌﾞｰﾄ照合OK")
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "OSﾌﾞｰﾄ読込み照合ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

        Exit Sub

OS1:    'OSﾌﾞｰﾄ1                    -5P、-9Zのみ
        Call WriteLog(False, "OSﾌﾞｰﾄ読込み")
        Call TestPreparation(2) : Call WriteLog(False, mLOG)                '前準備。RPのｱﾄﾞﾚｽ0x20000を0x20000に見えるようにする
        Try
            '読込み
            Call WriteLog(False, "0x27FFEﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 4, "00027FFE") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 4))
            mOSBTCalcCheckSum = ((mReadCheckSum >> 8) + (mReadCheckSum << 8) And &HFFFF) + (mReadCheckSum >> 8) + (mReadCheckSum And &HFF)
            WriteLog(False, "OSﾌﾞｰﾄCheckSUM:" & Convert.ToString(mOSBTCalcCheckSum, 16))
            '照合
            If mOSBTCalcCheckSum <> CInt("&H" & mFirmOBCheckSum) Then
                mErrorComment = "OSﾌﾞｰﾄCheckSUM照合異常。読込み:" & Convert.ToString(mOSBTCalcCheckSum, 16) & " FHSUM：" & mFirmOBCheckSum
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "OSﾌﾞｰﾄ照合OK")
        Catch ex As Exception
            cMcm = Nothing
            If mErrorComment = "" Then mErrorComment = "OSﾌﾞｰﾄ読込み照合ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        'OS読込み
        Call WriteLog(False, "OS読込み")
        Call TestPreparation(0) : Call WriteLog(False, mLOG)                '前準備。RPのｱﾄﾞﾚｽ0x00000を0x20000に見えるようにする
        Dim DAT1 As String
        Dim DAT2 As String
        Dim DAT3 As String
        Dim DAT4 As String

        Try
            '読込み
            Call WriteLog(False, "0x2DFE0ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002DFE0") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 4))
            DAT1 = Chr(mReadCheckSum And &HFF) & Chr(mReadCheckSum >> 8)
            Call WriteLog(False, "0x2DFE2ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002DFE2") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 4))
            DAT2 = Chr(mReadCheckSum And &HFF) & Chr(mReadCheckSum >> 8)
            Call WriteLog(False, "0x2DFE4ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002DFE4") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 4))
            DAT3 = Chr(mReadCheckSum And &HFF) & Chr(mReadCheckSum >> 8)
            Call WriteLog(False, "0x2DFE6ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002DFE6") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 4))
            DAT4 = Chr(mReadCheckSum And &HFF) & Chr(mReadCheckSum >> 8)
            Call WriteLog(False, DAT1 & DAT2 & DAT3 & DAT4)
            '照合
            If DAT1 & DAT2 & DAT3 & DAT4 <> mFirmOSCheckSum Then
                mErrorComment = "OSCheckSUM照合異常。読込み:" & DAT1 & DAT2 & DAT3 & DAT4 & " FHSUM：" & mFirmOSCheckSum
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "OS照合OK")
        Catch ex As Exception
            cMcm = Nothing
            mErrorComment = "OS読込み照合ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        If InStr(Model, "-5P") = 0 Then Exit Sub

        'OS読込み                   -5Pのみ
        Call WriteLog(False, "OS読込み")
        Call TestPreparation(1) : Call WriteLog(False, mLOG)                '前準備。
        
        Try
            '読込み
            Call WriteLog(False, "0x2DFE0ﾘｰﾄﾞ") : Call cMcm.IO_READ(0, 2, "0002DFE0") : Call WriteLog(False, "Recv:" & cMcm.RecvText)
            mReadCheckSum = CInt("&H" & Strings.Right(cMcm.RecvText, 2))
            '照合
            If mReadCheckSum = &HFFFF Then Throw New ApplicationException(mErrorComment)
        Catch ex As Exception
            cMcm = Nothing
            mErrorComment = "OS読込み照合ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try


    End Sub


    Public Sub I50(ByVal iName As String)
        '==============================================================================================================
        ' Function : I50:SLOT/XDFAIL TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim wTime As Double = 0

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call WriteLog(False, "ｽﾛｯﾄｱﾄﾞﾚｽ & XDFAILﾃｽﾄ中・・・約30秒")
        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(60)
            If cKP.SlotAddressTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄ異常:" & cKP.ERR_MSG
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "ｽﾛｯﾄｱﾄﾞﾚｽ & XDFAILﾃｽﾄOK")

    End Sub

    Public Sub I51(ByVal iName As String)
        '==============================================================================================================
        ' Function : I51:IOﾊﾞｽ-XACSﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call WriteLog(False, "XACS中・・・約2秒")
        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cKP.XACSTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XACSﾃｽﾄ異常:" & cKP.ERR_MSG
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "XACS検査OK")

    End Sub

    Public Sub I52(ByVal iName As String)
        '==============================================================================================================
        ' Function : I52:BCLR TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'XBCLRﾃｽﾄ
        Call WriteLog(False, "XBCLR検査中・・・約1秒")
        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cKP.XBCLRTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XBCLRﾃｽﾄ異常:" & cKP.ERR_MSG
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "XBCLR検査OK")

    End Sub

    Public Sub I53(ByVal iName As String)
        '==============================================================================================================
        ' Function : I53:IOﾊﾞｽ-XBRﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'XBRﾃｽﾄ
        Call WriteLog(False, "XBR検査中・・・約3秒")
        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cKP.XBRTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XBRﾃｽﾄ異常:" & cKP.ERR_MSG
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "XBR検査OK")

    End Sub


    Public Sub I55(ByVal iName As String)
        '==============================================================================================================
        ' Function : I55:IOﾊﾞｽ-割込みﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim mResStr As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call WriteLog(False, "XIRQ検査中・・・約5秒")

        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cKP.XIRQTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XIRQﾃｽﾄ異常:" & cKP.ERR_MSG
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "XIRQ検査OK")

    End Sub

    Public Sub I56(ByVal iName As String)
        '==============================================================================================================
        ' Function : I56:IOﾊﾞｽ-XLFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call WriteLog(False, "XLFAIL検査中・・・約1秒")

        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cKP.XLFAILTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XLFAILﾃｽﾄ異常:" & cKP.ERR_MSG
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "XLFAIL検査OK")

    End Sub

    Public Sub I57(ByVal iName As String)
        '==============================================================================================================
        ' Function : I57:IOﾊﾞｽ-XPFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call WriteLog(False, "XPFAIL検査中・・・約1秒")

        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cKP.XPFAILTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XPFAILﾃｽﾄ異常:" & cKP.ERR_MSG
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "XPFAIL検査OK")

    End Sub

    Public Sub I58(ByVal iName As String)
        '==============================================================================================================
        ' Function : I58:IOﾊﾞｽ-XSFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call WriteLog(False, "XSFAIL検査中・・・約1秒")

        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 5
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cKP.XSFAILTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XSFAILﾃｽﾄ異常:" & cKP.ERR_MSG
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "XSFAIL検査OK")

    End Sub

    Public Sub I59(ByVal iname As String)
        '==============================================================================================================
        ' Function : I59:IOﾊﾞｽ検査前準備
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim slot As Integer = 5

        Try
            Double.Parse(Strings.Right(iname, 1))
            slot = CInt(Strings.Right(iname, 1))        '文字列の一番右の数字がｽﾛｯﾄ番号
        Catch
            '文字列の一番右が数字でなければｽﾛｯﾄ5で前準備
        End Try

        Call WriteLog(False, "")
        Call WriteLog(False, iname & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Call WriteLog(False, "IOﾊﾞｽﾃｽﾄ前準備中・・・約5秒")

        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = slot
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(15)
            If cKP.TestPreparation() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽﾃｽﾄ前準備失敗"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cKP.LOG)
            If cKP.ERR_MSG <> "" Then Call WriteLog(False, cKP.ERR_MSG)
            cKP = Nothing
        End Try

        Call WriteLog(False, "IOﾊﾞｽﾃｽﾄ前準備OK")

    End Sub



    Public Sub I90(ByVal iName As String)
        '==============================================================================================================
        ' Function : I90:ﾕﾆｯﾄ電源ON
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim slot As Integer = 1

        Dim rCmd As String = ""

        If InStr(iName, "I90") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================

        Try
            Double.Parse(Strings.Right(iName, 1))
            slot = CInt(Strings.Right(iName, 1))        '文字列の一番右の数字がｽﾛｯﾄ番号
        Catch
            '文字列の一番右が数字でなければｽﾛｯﾄ2（ｱﾄﾞｵﾝCPU）で起動
        End Try

        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入：ｽﾛｯﾄ" & CStr(slot))

        Try
            cBUS = New cls122_FAM3_IOBUS_Ctrl
            Call cBUS.BUSMONITOR_TB04_CPU_START(slot)
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cBUS = Nothing
        End Try

        Call WaitStart(7)
        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源投入OK")

    End Sub

    Public Sub I91(ByVal iName As String)
        '==============================================================================================================
        ' Function : I91:ﾕﾆｯﾄ電源OFF
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        If InStr(iName, "I91") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================
        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源遮断")

        Try
            cBUS = New cls122_FAM3_IOBUS_Ctrl
            Call cBUS.BUSMONITOR_TB04_END()
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cBUS = Nothing
        End Try

        Call WriteLog(False, "検査ﾓｼﾞｭｰﾙ電源遮断OK")

    End Sub

    Public Sub I92(ByVal iName As String)
        '==============================================================================================================
        ' Function : I92:検査開始準備
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3RPxx\SetUp.ppt"
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                If dfrm.DIALOG_RESULT <> "" Then
                    Exit Do
                End If
                DoEvents()
            Loop

            If dfrm.DIALOG_RESULT = "NG" Then
                mErrorComment = "検査準備　異常"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            dfrm.Close()
            dfrm = Nothing

        End Try


    End Sub

    Public Sub I93(ByVal iName As String)
        '==============================================================================================================
        ' Function : I93:検査終了時メッセージ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        'Dim scmd As String
        Dim rCmd As String = ""

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Dim mReadLine As String = ""

        Dim temp(12) As String
        Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"

        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rData() As String = {}

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Try

            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3RPxx\Fin.ppt"
            'cSys.SignalTower("GREEN", "OFF", "ON", "OFF")
            dfrm.LED_COLOR = "YELLOW" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"
            dfrm.Button2.Visible = False
            dfrm.Button1.Left = CInt(dfrm.Width / 2) - CInt(dfrm.Button1.Width / 2)

            Call dfrm.frmShow()

            Do
                If dfrm.DIALOG_RESULT <> "" Then
                    Exit Do
                End If
                DoEvents()
            Loop

            If dfrm.DIALOG_RESULT = "NG" Then
                mErrorComment = "検査中止"
                dfrm.Close()
                Throw New ApplicationException(mErrorComment)
            End If

            dfrm.Close()
            dfrm = Nothing

        Catch ex As Exception

            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

            dfrm = Nothing

        End Try

        '最終確認
        Call WriteLog(False, "*")

        Call WriteLog(False, "●　検査合格時の作業  ======================================")
        Call WriteLog(False, "・主銘板に印字された内容を確認して、製品に貼り付ける。")
        Call WriteLog(False, "-以上、お疲れ様でした。(^o^) -")

    End Sub
    '------------------------------------------Ver 0.26  KSJ ADD Start -------------------------------------------------
    Public Sub I94(ByVal iName As String)
        '==============================================================================================================
        ' Function : I94:Result Log Save
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 検査合格の場合だけ作成
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Dim res As String() = {}
        'Make Folder ========================================================
        Dim SearchArea As String = "\\10.36.14.215\Fam3\result_LOG"
        Dim SearchYear As String = Format(DateTime.Now, "yyyy")
        Dim SearchMonth As String = Format(DateTime.Now, "yyyy-MM")
        Dim CalYear As String
        Dim CalMonth As String
        Dim CalTemp As String
        'Search date folder -----------------------------------------
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea
            cFop.FOLDER_NAME = SearchYear

            Call cFop.FOLDER_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cFop = Nothing

        End Try

        'Empty check ------------------------------------------------
        Dim sFolder As String = String.Empty
        For Each stFilePath As String In res
            sFolder &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then make folder -------------------------------------
        If sFolder = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.FOLDER_PATH = SearchArea & "\" & SearchYear

                Call cFop.FOLDER_MAKE()

            Catch ex As Exception
                mErrorComment = cFop.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString

            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try

            If mErrorComment <> "" Then Exit Sub
        End If

        'Make File ========================================================
        'Search month date file ------------------------------------------- 
        Try
            cFop = New cls012_File_Operation

            cFop.ROOT_FOLDER = SearchArea & "\" & SearchYear
            cFop.FILE_NAME = SearchMonth & ".mdb"

            Call cFop.FILE_SEARCH(res)

        Catch ex As Exception
            mErrorComment = cFop.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
        Finally
            mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

        End Try
        If mErrorComment <> "" Then Exit Sub

        'Empty check ------------------------------------------------
        Dim sFile = String.Empty
        sFile = String.Empty
        For Each stFilePath As String In res
            sFile &= stFilePath & System.Environment.NewLine
        Next stFilePath

        'Empty then copy file ---------------------------------------
        If sFile = String.Empty Then
            Try
                cFop = New cls012_File_Operation

                cFop.ORIGINAL_PATH = SearchArea & "\" & "ResultMaster.mdb"
                cFop.FOLDER_PATH = SearchArea & "\" & SearchYear & "\"     '& SearchFolder & "\"
                cFop.FILE_NAME = SearchMonth & ".mdb"

                Call cFop.FILE_COPY()

            Catch ex As Exception
                mErrorComment = "ResultMaster.mdbがありません。"
                Throw New ApplicationException(mErrorComment)
            Finally
                mErrorComment = cFop.ERR_COMMENT : cFop = Nothing

            End Try
            If mErrorComment <> "" Then Exit Sub

        End If

        'Result Dataの読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String = "tblResult"
        Dim rField As String = "StartDateTime, MSCODE, SerialNo, EquipmentNo, RunMode"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim fNameDel() As String = {""}
        Dim sDataDel() As String = {""}

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

        'Search and delete redundant data=======================================================
        If sData(4) = "MAN" Then
            Call WriteLog(False, "Manual検査時には Reuslt Log Saveしません。")
            Exit Sub
        ElseIf sData(4) = "AUT" Then
            For i = 0 To 1
                dbName = SearchArea & "\" & SearchYear & "\" & SearchMonth & ".mdb"

                Select Case i 'yyyy/mm 再設定
                    Case 0 '現月の場合
                    Case 1 '先月の場合
                        If Strings.Right(SearchMonth, 2) = "01" Then ' たとえば、2018/01なら 2017/12で検索 

                            CalYear = CStr(CInt(SearchYear) - 1)
                            CalMonth = "12"
                            dbName = SearchArea & "\" & CalYear & "\" & CalYear & "-" & CalMonth & ".mdb"

                        Else ' 2月～12月なら -1ヵ月

                            CalTemp = CStr(CInt(Strings.Right(SearchMonth, 2)) - 1)

                            If CInt(CalTemp) >= 10 Then
                                CalMonth = SearchYear & "-" & CalTemp
                            Else
                                CalMonth = SearchYear & "-0" & CalTemp
                            End If

                            dbName = SearchArea & "\" & SearchYear & "\" & CalMonth & ".mdb"

                        End If
                End Select

                Try
                    cADO = New cls010_MDB_Operation
                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName

                    sSQL = "SELECT * FROM " & tbName & " WHERE MSCODE = '" & sData(1) & "' AND SerialNo = '" & sData(2) & "'"
                    cADO.DB_SQL = sSQL

                    Call cADO.DB_READ_JET(fNameDel, sDataDel)

                    If cADO.DB_DATA_COUNT > 0 Then
                        sSQL = "DELETE FROM " & tbName & " WHERE MSCODE = '" & sDataDel(0) & "' AND SerialNo = '" & sDataDel(1) & "'"
                        cADO.DB_PATH = dbName
                        cADO.TB_NAME = tbName
                        cADO.DB_SQL = sSQL

                        Call cADO.DB_EXEC_JET()
                    End If
                Catch ex As Exception
                    If mErrorComment = "" Then mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                Finally
                    cADO = Nothing
                End Try
            Next
        Else
            mErrorComment = "Run Mode 異常"
            Throw New ApplicationException(mErrorComment)

        End If
        'Write of Data ---------------------------------------------
        dbName = SearchArea & "\" & SearchYear & "\" & SearchMonth & ".mdb"
        tbName = "tblResult"

        sSQL = "INSERT INTO " & tbName & " ("
        sSQL = sSQL & fName(0) & ", "
        sSQL = sSQL & fName(1) & ", "
        sSQL = sSQL & fName(2) & ", "
        sSQL = sSQL & fName(3)
        sSQL = sSQL & ") VALUES("
        sSQL = sSQL & "'" & sData(0) & "', "
        sSQL = sSQL & "'" & sData(1) & "', "
        sSQL = sSQL & "'" & sData(2) & "', "
        sSQL = sSQL & "'" & sData(3) & "'"
        sSQL = sSQL & ")"

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing
        End Try

    End Sub
    '------------------------------------------Ver 0.26 KSJ ADD End -------------------------------------------------
    Private Function IA0() As String
        '==============================================================================================================
        ' Function : IA0:LineSerial
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        IA0 = ""

        Dim wTime As Double = 0

        Dim LSerial As String
        '==========================================================================================

        Try
            'Get Line Serial -----------------------------
            cSer = New cls030_LineSerial

            Call cSer.IssueLineSerial()
            If cSer.ERR_MSG <> "" Then
                mErrorComment = cSer.ERR_MSG
                Throw New ApplicationException(mErrorComment)
            End If

            LSerial = cSer.LineSerial

            cSer = Nothing

            Call WriteData("LineSerial", LSerial)
            Call WriteLog(False, "LineSerial : " & LSerial)

            ' ==============================================================

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

        End Try

        IA0 = LSerial

    End Function

    Public Sub IP1(ByVal iName As String)
        '==============================================================================================================
        ' Function : IP1:PrintLabel
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "StartDatetime"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            rDTkey = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

        Call LabelPrint(dbName, rDTkey)

        '主銘板再印字ボタン有効化 -=========================================================
        Insfrm.btnRePrintLabel.Enabled = True
        AddHandler Insfrm.btnRePrintLabel.Click, AddressOf RePrint_Click

    End Sub

    Public Sub IP2(ByVal iName As String)
        '==============================================================================================================
        ' Function : IP2: OK PrintLabel
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : RP6x,SP69-9S/C12用OKのみのﾗﾍﾞﾙﾌﾟﾘﾝﾄ
        '==============================================================================================================

        Dim Model As String = ModelRead()
        Dim tbName As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim rField As String

        Dim mPrinterName As String

        ' ==============================================================
        'ラベルプリンタ名の読込み
        cADO = New cls010_MDB_Operation

        cADO.DB_PATH = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        tbName = "tblEquipmentSetting"
        cADO.TB_NAME = tbName
        rField = "LabelPrinter"

        sSQL = "SELECT " & rField & " FROM " & tbName
        cADO.DB_SQL = sSQL

        Call cADO.DB_READ_JET(fName, sData)
        If cADO.DB_DATA_COUNT = 0 Then
            mErrorComment = "印字するプリンタ名を取得できませんでした。"
            Throw New ApplicationException(mErrorComment)
        End If

        mPrinterName = sData(0)

        Try
            cLbl = New cls011_Excel_Operation
            With cLbl
                'Select Case Microsoft.VisualBasic.Left(sModel, 6)
                '    Case "F3UM01" : .FILE_PATH = mAppPath & "410 Print Form\" & "F3UM01Label.xls"
                '    Case Else
                '        mErrorComment = "主銘板を印字するExcelファイルが存在ません。"
                '        Throw New ApplicationException(mErrorComment)
                'End Select
                .FILE_PATH = mAppPath & "410 Print Form\" & "OKLabel.xls"
                .SHEET_NAME = "Sheet1"
                .EXCEL_START()

                .CELL_DATA_SET("MODEL", Model)
                .CELL_DATA_SET("DATE", Now.ToString("yyyy/MM/dd"))

                .CELL_DATA_SET("MODELH", "OK")
                .CELL_DATA_SET("BARCODE", "OK")
                '.BARCODE_SHOW("BARCODE", 120 + 190 - 120, 260, 120, 27)
                .BARCODE_SHOW("BARCODE", 120 + 190 - 120, 285, 120, 27)

                .PRINT_EXCEL(mPrinterName)

            End With
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLbl.CLOSE_EXCEL()
            cLbl = Nothing
        End Try

        Try

            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3RPxx\Fin.ppt"
            'cSys.SignalTower("GREEN", "OFF", "ON", "OFF")
            dfrm.LED_COLOR = "YELLOW" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"
            dfrm.Button2.Visible = False
            dfrm.Button1.Left = CInt(dfrm.Width / 2) - CInt(dfrm.Button1.Width / 2)

            Call dfrm.frmShow()

            Do
                If dfrm.DIALOG_RESULT <> "" Then
                    Exit Do
                End If
                DoEvents()
            Loop

            If dfrm.DIALOG_RESULT = "NG" Then
                mErrorComment = "検査中止"
                dfrm.Close()
                Throw New ApplicationException(mErrorComment)
            End If

            dfrm.Close()
            dfrm = Nothing

        Catch ex As Exception

            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Debug.Print(mErrorComment)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally

            dfrm = Nothing

        End Try

        '最終確認
        Call WriteLog(False, "*")

        Call WriteLog(False, "●　検査合格時の作業  ======================================")
        Call WriteLog(False, "後検査を行ってください。")
        
    End Sub

    Public Sub LabelPrint(ByVal dbName As String, ByVal rDTkey As String)

        'v0.02新主銘板対応
        Dim cLblPrnt As cls034_LabelPrint

        cLblPrnt = New cls034_LabelPrint
        Try
            cLblPrnt.PrintForm = "RPxx"
            Call cLblPrnt.LabelPrint(dbName, rDTkey)
        Catch ex As Exception
            mErrorComment = cLblPrnt.ERR_COMMENT
            If mErrorComment = "" Then
                mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLblPrnt = Nothing
        End Try

    End Sub

    Private Sub RePrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : 主銘板再印字 
        'Remarks  :
        '==============================================================================================================

        Call IP1("主銘板再印字")

    End Sub

    Private Sub TestPreparation(ByVal add As Integer)

        Dim cMcm1 As cls120_SP_M_Command

        Dim Adres1 As String = "0000009A"
        Dim Adres2 As String = "000000F8"
        Dim Adres3 As String = "0000008E"
        Dim Adres4 As String = "00000090"
        Dim Adres5 As String = "00000092"

        Dim writeDATA_0x90 As String = ""
        Dim writeDATA_0x92 As String = ""

        mLOG = ""
        mLOG &= "IOﾊﾞｽ経由読込み前準備" & vbCrLf

        Select Case add
            Case 0      'IO_READで0x20000番地にRPの0番地が見える
                writeDATA_0x90 = "0000" : writeDATA_0x92 = "0010"
            Case 1
                writeDATA_0x90 = "0008" : writeDATA_0x92 = "0010"
            Case 2      'IO_READで0x20000番地にRPの0x20000番地が見える
                writeDATA_0x90 = "0000" : writeDATA_0x92 = "8010"
            Case 7      'IO_READで0x20000番地にRPの0x70000番地が見える
                writeDATA_0x90 = "0004" : writeDATA_0x92 = "C010"
        End Select

        cMcm1 = New cls120_SP_M_Command
        cMcm1.IPAdd = IP_U2                      'CPU用LE経由
        cMcm1.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm1.SlotNo = 1                         'CPU Slot番号
        cMcm1.UnitNo = 0                         'ユニット番号

        Try
            Call cMcm1.IO_WRITE(1, 2, Adres1, "8080")        'IOREGｱﾄﾞﾚｽ0x9aに0x8080をﾗｲﾄ
            mLOG &= "Send : " & "IOREGｱﾄﾞﾚｽ0x9A（ｽﾛｯﾄ2）に0x8080をﾗｲﾄ" & vbCrLf
            mLOG &= "Recv : " & cMcm1.RecvText & vbCrLf
            Call cMcm1.IO_WRITE(1, 2, Adres2, "0000")       'IOREGｱﾄﾞﾚｽ0xF8に0x0000をﾗｲﾄ
            mLOG &= "Send : " & "IOREGｱﾄﾞﾚｽ0xF8（ｽﾛｯﾄ2）に0x0000をﾗｲﾄ" & vbCrLf
            mLOG &= "Recv : " & cMcm1.RecvText & vbCrLf
            Call cMcm1.IO_WRITE(1, 2, Adres3, "8012")        'IOREGｱﾄﾞﾚｽ0x8eに0x8012をﾗｲﾄ
            mLOG &= "Send : " & "IOREGｱﾄﾞﾚｽ0x8E（ｽﾛｯﾄ2）に0x8012をﾗｲﾄ" & vbCrLf
            mLOG &= "Recv : " & cMcm1.RecvText & vbCrLf
            Call cMcm1.IO_WRITE(1, 2, Adres4, writeDATA_0x90)        'IOREGｱﾄﾞﾚｽ0x90にwriteDATA_0x90をﾗｲﾄ
            mLOG &= "Send : " & "IOREGｱﾄﾞﾚｽ0x90（ｽﾛｯﾄ2）に0x" & writeDATA_0x90 & "をﾗｲﾄ" & vbCrLf
            mLOG &= "Recv : " & cMcm1.RecvText & vbCrLf
            Call cMcm1.IO_WRITE(1, 2, Adres5, writeDATA_0x92)        'IOREGｱﾄﾞﾚｽ0x92にwriteDATA_0x92をﾗｲﾄ
            mLOG &= "Send : " & "IOREGｱﾄﾞﾚｽ0x92（ｽﾛｯﾄ2）に0x" & writeDATA_0x92 & "をﾗｲﾄ" & vbCrLf
            mLOG &= "Recv : " & cMcm1.RecvText & vbCrLf

        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽ経由ﾃﾞｰﾀ読込み前準備ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm1 = Nothing
        End Try

        Call mSecWait(100)

    End Sub

    Private Sub WaitStart(ByVal sVal As Double)
        '==========================================================================================
        ' Function : Wait Start
        ' Data     : Setting time
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Call Insfrm.WaitStart(sVal)

    End Sub

    Private Sub WaitStop()
        '==========================================================================================
        ' Function : Wait Stop
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Call Insfrm.WaitStop()

    End Sub

    Private Sub CountStart(ByVal sVal As Double)
        '==========================================================================================
        ' Function : Count Start
        ' Data     : Setting time
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Call Insfrm.CountStart(sVal)

    End Sub

    Private Sub CountStop()
        '==========================================================================================
        ' Function : Count Stop
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Call Insfrm.CountStop()

    End Sub

    Private Sub InspectionTimerStart()
        '==========================================================================================
        ' Function : Inspection timer start
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Insfrm.InspectionTimerStart()

    End Sub

    Private Sub InspectionTimerStop()
        '==========================================================================================
        ' Function : Inspection timer Stop
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Insfrm.InspectionTimerStop()

    End Sub

    Private Sub WriteData(ByVal fName As String, ByVal sData As String)
        '==========================================================================================
        ' Function : Write data(Temporary.mdb)
        ' Data     : Field name , data
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""

        dbName = mAppPath & "210 Result Data\Temporary.mdb"
        tbName = "tblResult"

        'Write of data ----------------------------------------------
        sSQL = "UPDATE " & tbName & " SET "
        sSQL &= fName & " = '" & sData & "'"

        Try
            cADO = New cls010_MDB_Operation

            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_EXEC_JET()

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cADO = Nothing

        End Try

    End Sub

    Private Sub WriteLog(ByVal TmStmp As Boolean, ByVal msg As String)
        '==========================================================================================
        ' Function : Write Log(Temporary.txt)
        ' Data     : Time stamp Y/N , message
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim lFile As String = mAppPath & "220 Result Log\Temporary.txt"
        Try
            cTxt = New cls013_TextFile_Operation
            cTxt.FILE_PATH = lFile

            If msg <> "*" Then
                cTxt.WRITE_TEXT_FILE(TmStmp, msg)
            Else
                cTxt.WRITE_TEXT_LINE(TmStmp, "*", 50)
            End If

        Catch ex As Exception
            mErrorComment = cTxt.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString

        Finally
            cTxt = Nothing

        End Try
        DoEvents()

        Call Insfrm.LogBrowserNavi()

    End Sub

    Private Sub WaitTimer(ByVal tm As Double)
        '==========================================================================================
        ' Function : Wait timer
        ' Data     : tm : Wait time(Sec)
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim sTime As Double
        Dim nTime As Double
        Dim eTime As Double

        sTime = Microsoft.VisualBasic.DateAndTime.Timer()

        Do While tm > eTime
            nTime = Microsoft.VisualBasic.DateAndTime.Timer()
            eTime = nTime - sTime
            If eTime < 0 Then eTime = 24.0# * 3600.0# - sTime + nTime
            DoEvents()
        Loop

        Debug.Print("Wait:" & CStr(eTime))

    End Sub

    Private Function SerialRead() As String

        'PLASMA Serial番号の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "SerialNo"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        SerialRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            SerialRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function

    Private Function ModelRead() As String

        '検査型式の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "InspectionModel"            'InspectionModel  MSCODE
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        ModelRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            ModelRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
    End Sub

End Class
