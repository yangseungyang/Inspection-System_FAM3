Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm020_InspectionStartEnd
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.Net.NetworkInformation

Public Class cls2D8_F3LE1x_inspection
    'Class ================================================
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cLnk As cls080_F3LinkToLE
    Dim cDMM As cls101_ComRS232C
    Dim cISPComm As cls101_ComRS232C
    Dim cWin As cls001_Window_Control
    Dim cMac As cls031_MacAddress
  
    Dim cSer As cls030_LineSerial
    Dim cLbl As cls011_Excel_Operation
    Dim cMcm As cls120_SP_M_Command
    Dim cPCSP As cls060_PCNetSPEED
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl
    Dim cTEST As cls121_SP_TEST_MODE
    Dim cKP As cls124_KP_IOBUS_TEST
    Dim cSPKP As cls2D6_F3SPxx_KP_Inspection

    Dim dfrm As frm500_MessagePopUp

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String

    'Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3のLE
    Dim IP_U2 As String = "192.168.1.11"        'ﾒｲﾝCPU
    Dim IP_LE As String = "192.168.0.2"        '被検査ﾓｼﾞｭｰﾙ

    Dim mTestComAddress As Integer = &H236      'ｽﾛｯﾄ5で検査$800036
    Dim mTestExeAddress As Integer = &H234      'ｽﾛｯﾄ5で検査$800034
    Dim mTestResAddress As Integer = &H23C      'ｽﾛｯﾄ5で検査$80003C

    Dim mTimeOut As Integer

    Dim iName As String

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
            'Case "I01" : Call I01(iName)            'I01:接続準備
            Case "I02" : Call I02(iName)            'I02:ﾌｧｰﾑｳｪｱ書き込み
            Case "I03" : Call I03(iName)            'I03:ﾓｼﾞｭｰﾙIDﾁｪｯｸ
            Case "I04" : Call I04(iName)            'I04:ROMSUMﾁｪｯｸ
            Case "I05" : Call I05(iName)            'I05:SRAM READ/WRITE TEST
            Case "I06" : Call I06(iName)            'I06:SDRAM READ/WRITE TEST
            Case "I07" : Call I07(iName)            'I07:CLOCK TEST
            Case "I08" : Call I08(iName)            'I08:LAN IRQ TEST
            Case "I09" : Call I09(iName)            'I09:XSFAIL TEST
            Case "I10" : Call I10(iName)            'I10:EXT TIME CHECK
            Case "I11" : Call I11(iName)            'I11:10msec TIME IRQ CHECK
            Case "I12" : Call I12(iName)            'I12:CPLD CHECKﾊﾞﾝｸ切り替え確認
            Case "I13" : Call I13(iName)            'I13:ﾛｰﾀﾘｰｽｲｯﾁ検査
            Case "I14" : Call I14(iName)            'I14:ﾃﾞｨｯﾌﾟｽｲｯﾁ検査
            Case "I15" : Call I15(iName)            'I15:MAC ID READ
            Case "I16" : Call I16(iName)            'I16:ﾛｰﾀﾘｰｽｲｯﾁ通信設定
            Case "I17" : Call I17(iName)            'I17:出荷設定
                'Case "I18" : Call I18(iName)            'I18:FG導通ﾃｽﾄ
                'Case "I19" : Call I19(iName)            'I19:SHIELD導通ﾃｽﾄ
                'Case "I20" : Call I20(iName)            'I20:電源電圧ﾃｽﾄ
            Case "I21" : Call I21(iName)            'I21:LED検査

            Case "I30" : Call I30(iName)            'I30:Ethernet通信ﾃｽﾄ

            Case "I50" : Call I50(iName)            'I50:SLOT/XDFAIL TEST
                'Case "I51" : Call I51(iName)            'I51:ACS SIGNAL TEST
            Case "I52" : Call I52(iName)            'I52:BCLR TEST
            Case "I53" : Call I53(iName)            'I53:BR TEST
                'Case "I54" : Call I54(iName)            'I54:I/O BUS WRITE/READ TEST
            Case "I55" : Call I55(iName)            'I55:I/O BUS INTERRUPT TEST
            Case "I56" : Call I56(iName)            'I56:XLFAIL TESTｽﾛｯﾄ5
            Case "I57" : Call I57(iName)            'I57:XPFAIL TEST
            Case "I58" : Call I58(iName)            'I58:XSFAIL TEST
            Case "I59" : Call I59(iName)            'I59:IOﾊﾞｽ検査前準備

            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF



                'Case "I92" : Call I92(iName)            'I92:検査開始時処理
                'Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save        'Rev 0.26 KSJ ADD
            Case "IA1" : Call IA1(iName)            'IA1:MACｱﾄﾞﾚｽ採番
            Case "IA2" : Call IA2(iName)            'IA2:MACｱﾄﾞﾚｽ重複確認

            Case "IP1" : Call IP1(iName)            'IP1:PrintLabel

                'Case Else : mErrorComment = "検査プロシージャが定義されていません。"   'Rev0.11DEL
            Case Else : mErrorComment = "검사 프로시져가 정의 되어있지 않습니다."   'Rev0.11ADD
                MsgBox(mErrorComment)

        End Select


    End Sub

    Public Sub I02(ByVal iName As String)
        '==============================================================================================================
        ' Function : I02:ﾌｧｰﾑｳｪｱ書き込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I02(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I03(ByVal iName As String)
        '==============================================================================================================
        ' Function : I03:ﾓｼﾞｭｰﾙIDﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim RecvDataLE As String = ""
        Dim Model As String = ModelRead()

        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2
        Call cLnk.F3PCLinkCommunication("01INF3,0", RecvDataLE)
        Debug.Print(RecvDataLE)
        Call WriteLog(False, "ｽﾛｯﾄ5読み取り：" & Mid(RecvDataLE, 5 + 8 * 4, 4))

        'ｽﾛｯﾄ5だけを見る
        If Mid(RecvDataLE, 5 + 8 * 4, 4) <> Mid(Model, 3, 4) Then
            mErrorComment = "ﾓｼﾞｭｰﾙ識別異常：" & Mid(Model, 3, 4)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "ﾓｼﾞｭｰﾙ識別OK：" & Mid(Model, 3, 4))
        cLnk = Nothing

    End Sub

    Public Sub I04(ByVal iName As String)
        '==============================================================================================================
        ' Function : I04:ROMｻﾑﾁｪｯｸ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Dim Model As String = ModelRead()
        Dim dbName As String
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim ROM_SUM As String = ""

        'チェックサムの読込み
        rField = "RomSum"
        Try
            cADO = New cls010_MDB_Operation
            dbName = mAppPath & "400 Common File\MSCODE.mdb"
            cADO.DB_PATH = dbName
            tbName = "tblMSCODE"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model ='" & Model & "'"
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)
            ROM_SUM = sData(0)
        Catch ex As Exception
            cADO = Nothing
        End Try

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0400")
        cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0400をﾗｲﾄ")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")

        If FlagCheck() = True Then
            cMcm.IO_READ(0, 2, mTestResAddress.ToString("X8"))
        Else
            mErrorComment = "ROMSUMﾁｪｯｸFlagﾁｪｯｸ異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        If Strings.Right(cMcm.RecvText, 4) <> ROM_SUM Then
            mErrorComment = "ROMSUMﾁｪｯｸ異常。期待値：" & ROM_SUM & "　戻り値：" & Strings.Right(cMcm.RecvText, 4)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "期待値：" & ROM_SUM & "　読み値：" & Strings.Right(cMcm.RecvText, 4))
        Call WriteLog(False, "ROMSUMﾁｪｯｸOK")
        cMcm = Nothing

    End Sub

    Public Sub I05(ByVal iName As String)
        '==============================================================================================================
        ' Function : I05:SRAM READ/WRITE TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0500")
        cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0500をﾗｲﾄ")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")

        If FlagCheck() = True Then

        Else
            mErrorComment = "SRAM READ/WRITE TEST異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "SRAM READ/WRITE TEST OK")
        cMcm = Nothing

    End Sub

    Public Sub I06(ByVal iName As String)
        '==============================================================================================================
        ' Function : I06:SDRAM READ/WRITE TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        mTimeOut = 100
        cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0600")
        cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0800をﾗｲﾄ")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")

        Call WriteLog(False, "約60秒お待ちください。")
        Try
            Call Insfrm.CountStart(100)
            If FlagCheck() = True Then

            Else
                mErrorComment = "SDRAM READ/WRITE TEST異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "SDRAM READ/WRITE TEST OK")
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
            Call Insfrm.CountStop()
        End Try
        

    End Sub

    Public Sub I07(ByVal iName As String)
        '==============================================================================================================
        ' Function : I07:CLOCK TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Dim mRecData As Integer

        Try
            cMcm = New cls120_SP_M_Command
            cMcm.IPAdd = IP_U2
            cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
            cMcm.SlotNo = 1                         'CPU Slot番号
            cMcm.UnitNo = 0                         'ユニット番号

            cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0700")
            cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0700をﾗｲﾄ")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")

            If FlagCheck() = True Then
                cMcm.IO_READ(0, 2, mTestResAddress.ToString("X8"))
                mRecData = Convert.ToInt32(Strings.Right(cMcm.RecvText, 4), 16)
                Call WriteLog(False, "ClockRead:" & mRecData.ToString("X4"))
            Else
                mErrorComment = "CLOCK TEST異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Select Case Strings.Left(Model, 9)
                Case "F3LE12-0T"
                    If mRecData < &H805 Or mRecData > &H815 Then
                        mErrorComment = "CLOCK TEST異常。規格：&H805～815の範囲外"
                        Throw New ApplicationException(mErrorComment)
                    End If
                Case "F3LE11-0T", "F3LE01-0T", "F3NX01-1N"
                    If mRecData < &H7B0 Or mRecData > &H7D0 Then
                        mErrorComment = "CLOCK TEST異常。規格：&H7B0～7D0の範囲外"
                        Throw New ApplicationException(mErrorComment)
                    End If
            End Select

            Call WriteLog(False, "CLOCK TEST OK")

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

    End Sub

    Public Sub I08(ByVal iName As String)
        '==============================================================================================================
        ' Function : I08:LAN IRQ TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'ｳｨﾝﾄﾞｳｱﾄﾞﾚｽを変更
        Call cMcm.IO_WRITE(1, 2, "00000211", "0060")
        Call cMcm.IO_WRITE(1, 2, "00000212", "2C38")
        Call cMcm.IO_WRITE(0, 2, "00000225", "0001")
        Call cMcm.IO_WRITE(0, 2, "00000227", "0002")

        'ｳｨﾝﾄﾞｳｱﾄﾞﾚｽを戻す
        Call cMcm.IO_WRITE(1, 2, "00000211", "0080")
        Call cMcm.IO_WRITE(1, 2, "00000212", "0010")

        cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0800")
        cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0800をﾗｲﾄ")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")

        If FlagCheck() = True Then

        Else
            mErrorComment = "LAN IRQ TEST異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "LAN IRQ TEST OK")
        cMcm = Nothing

    End Sub

    Public Sub I09(ByVal iName As String)
        '==============================================================================================================
        ' Function : I09:XSFAIL TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0900")
        cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0900をﾗｲﾄ")
        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")

        If FlagCheck() = True Then

        Else
            mErrorComment = "XSFAIL TEST異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "XSFAIL TEST OK")
        cMcm = Nothing

    End Sub

    Public Sub I10(ByVal iName As String)
        '==============================================================================================================
        ' Function : I10:EXT TIME CHECK
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim mRecData As Integer

        Try
            cMcm = New cls120_SP_M_Command
            cMcm.IPAdd = IP_U2
            cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
            cMcm.SlotNo = 1                         'CPU Slot番号
            cMcm.UnitNo = 0                         'ユニット番号

            cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0A00")
            cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0A00をﾗｲﾄ")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")
            If FlagCheck() = True Then
                cMcm.IO_READ(0, 2, mTestResAddress.ToString("X8"))
                mRecData = Convert.ToInt32(Strings.Right(cMcm.RecvText, 4), 16)
                Call WriteLog(False, "Read:" & mRecData.ToString("X4"))
            Else
                mErrorComment = "EXT TIME CHECK異常"
                Throw New ApplicationException(mErrorComment)
            End If

            If mRecData < &HA Or mRecData > &HC Then
                mErrorComment = "規格範囲外で異常。&HA～C"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "EXT TIME CHECK OK")

        Catch
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

    End Sub

    Public Sub I11(ByVal iName As String)
        '==============================================================================================================
        ' Function : I11:10msec TIME IRQ CHECK
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim mRecData As Integer

        Try
            cMcm = New cls120_SP_M_Command
            cMcm.IPAdd = IP_U2
            cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
            cMcm.SlotNo = 1                         'CPU Slot番号
            cMcm.UnitNo = 0                         'ユニット番号

            cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0B00")
            cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0B00をﾗｲﾄ")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")
            If FlagCheck() = True Then
                cMcm.IO_READ(0, 2, mTestResAddress.ToString("X8"))
                mRecData = Convert.ToInt32(Strings.Right(cMcm.RecvText, 4), 16)
                Call WriteLog(False, "Read:" & mRecData.ToString("X4"))
            Else
                mErrorComment = "10msec TIME IRQ CHECK異常"
                Throw New ApplicationException(mErrorComment)
            End If

            If mRecData < &HA Or mRecData > &HC Then
                mErrorComment = "規格範囲外で異常。&HA～C"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "10msec TIME IRQ CHECK OK")

        Catch
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

    End Sub

    Public Sub I12(ByVal iName As String)
        '==============================================================================================================
        ' Function : I12:CPLD CHECKﾊﾞﾝｸ切り替え確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Dim mEffect As Boolean = True
        Dim mFuncName As String = ""
        Dim mStart As Integer
        Dim ddt1, ddt2, ddt3 As Integer
        ddt1 = 0 : ddt2 = 1 : ddt3 = 2

        Select Case Mid(Model, 1, 9)
            Case "F3LE01-0T", "F3NX01-1N"
                mStart = 32 : mEffect = False : mFuncName = "CPLD CHECKﾊﾞﾝｸ切り替え無効確認"
            Case Else
                mStart = 1 : mFuncName = "CPLD CHECKﾊﾞﾝｸ切り替え有効確認"
        End Select

        Call WriteLog(False, mFuncName)

        Try
            cMcm = New cls120_SP_M_Command
            cMcm.IPAdd = IP_U2
            cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
            cMcm.SlotNo = 1                         'CPU Slot番号
            cMcm.UnitNo = 0                         'ユニット番号

            cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0C00")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0C00をﾗｲﾄ")
            'Call WriteLog(False, "約30秒お待ちください。")
            Call Insfrm.CountStart(60)

            For i = 1 To 32
                Call cMcm.IO_WRITE(1, 2, "00000212", "0DF0")                '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定$803780を$200でｱｸｾｽできるように設定
                cMcm.IO_WRITE(0, 2, "00000200", ddt1.ToString("X4"))        '$803780    "X4" 16進に変換して"000*"4桁文字列表記にする
                Call cMcm.IO_WRITE(1, 2, "00000212", "0E10")                '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定$803800を$200でｱｸｾｽできるように設定
                cMcm.IO_WRITE(0, 2, "00000200", ddt2.ToString("X4"))        '$803800
                Call cMcm.IO_WRITE(1, 2, "00000212", "0FF0")                '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ設定$803F80を$200でｱｸｾｽできるように設定
                cMcm.IO_WRITE(0, 2, "0000027E", ddt3.ToString("X4"))        '$803FFE
                ddt1 += 1 : ddt2 += 1 : ddt3 += 1
            Next
            Call cMcm.IO_WRITE(1, 2, "00000212", "0010")                    '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽを元に戻す

            cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
            Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")
            If FlagCheck() = True Then
                Dim MADDAT As Integer
                For j = 0 To 1
                    Select Case j
                        Case 0 : MADDAT = &H23C
                        Case 1 : MADDAT = &H23E
                    End Select
                    For i = mStart To 32        '有効のとき1～無効のとき32だけ
                        If i = 18 Then
                            Call cMcm.IO_WRITE(1, 2, "00000212", "0030")    '$80を超えるとき1回だけ下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽをﾁｪﾝｼﾞする
                            MADDAT = MADDAT - &H80
                        End If
                        Debug.Print(MADDAT.ToString("X8"))
                        cMcm.IO_READ(0, 2, MADDAT.ToString("X8"))
                        Select Case j
                            Case 0
                                If Convert.ToInt32(Strings.Right(cMcm.RecvText, 4), 16) <> i Then
                                    mErrorComment = mFuncName & "異常"
                                    Throw New ApplicationException(mErrorComment)
                                End If
                            Case 1
                                If Convert.ToInt32(Strings.Right(cMcm.RecvText, 4), 16) <> i + 1 Then
                                    mErrorComment = mFuncName & "異常"
                                    Throw New ApplicationException(mErrorComment)
                                End If
                        End Select
                        MADDAT += 4
                    Next
                    Call cMcm.IO_WRITE(1, 2, "00000212", "0010")            '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽを戻す
                Next
            Else
                mErrorComment = mFuncName & "異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, mFuncName & " OK")

        Catch
            If mErrorComment = "" Then mErrorComment = mFuncName & "異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
            Call Insfrm.CountStop()
        End Try

    End Sub

    Public Sub I13(ByVal iName As String)
        '==============================================================================================================
        ' Function : I13:ﾛｰﾀﾘｰｽｲｯﾁ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        'Dim mMessage As String = "ﾛｰﾀﾘｰｽｲｯﾁSWを回してください。↑が0です。左上からSW1～"    'Rev0.11DEL
        Dim mMessage As String = "로타리SW를 돌려주세요. ↑가 0입니다. 좌상부터 SW1~ "    'Rev0.11DEL

        'メッセージ表示
        dfrm = New frm500_MessagePopUp
        'dfrm.DISP_TITLENAME = "ﾛｰﾀﾘｰｽｲｯﾁ検査"    'Rev0.11DEL
        dfrm.DISP_TITLENAME = "로타리SW 검사"    'Rev0.11ADD
        dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3LExx\RSW1.ppt"
        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        Call dfrm.frmShow()

        Try
            For i = 1 To 8
                mMessage &= vbCrLf & "SW" & i.ToString & "-"
                For ii = 1 To 15
                    Do
                        If Mid(RTSRead(), i, 1) = Convert.ToString(ii, 16).ToUpper Then Exit Do
                        DoEvents()
                        'wait push button -------------------
                        If dfrm.DIALOG_RESULT = "NG" Then
                            dfrm.Close()
                            dfrm = Nothing
                            mErrorComment = "ﾛｰﾀﾘｰｽｲｯﾁ検査 異常 "
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Loop

                    Beep()

                    If ii = 1 Then
                        dfrm.Close()
                        dfrm = New frm500_MessagePopUp
                        'mMessage &= "SW" & i.ToString & "ﾃｽﾄ中"   'Rev0.11DEL
                        mMessage &= "SW" & i.ToString & "TEST中"   'Rev0.11ADD
                        dfrm.DISP_MESSAGE = mMessage
                        dfrm.frmShow()
                    End If
                    dfrm.TextBox1.Text = mMessage & ii.ToString("X")
                Next

                Do
                    If RTSRead() = "00000000" Then Exit Do
                    DoEvents()
                    'wait push button -------------------
                    If dfrm.DIALOG_RESULT = "NG" Then
                        dfrm.Close()
                        dfrm = Nothing
                        mErrorComment = "ﾛｰﾀﾘｰｽｲｯﾁ検査 異常 "
                        Throw New ApplicationException(mErrorComment)
                    End If
                Loop

                dfrm.Close()
                dfrm = New frm500_MessagePopUp
                mMessage &= "SW" & i.ToString("X") & "OK"
                dfrm.DISP_MESSAGE = mMessage
                dfrm.frmShow()

            Next
        Catch
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "ﾛｰﾀﾘｰｽｲｯﾁ検査OK")

    End Sub

    Public Sub I14(ByVal iName As String)
        '==============================================================================================================
        ' Function : I13:ﾃﾞｨｯﾌﾟｽｲｯﾁ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        'Dim mMessage As String = "ﾃﾞｨｯﾌﾟｽｲｯﾁSWをON→OFFしてください。左から1->8" & vbCrLf    'Rev0.11DEL
        Dim mMessage As String = "DIP SW를 ON→OFF해 주세요. 좌부터1→8" & vbCrLf    'Rev0.11ADD

        'メッセージ表示
        dfrm = New frm500_MessagePopUp
        'dfrm.DISP_TITLENAME = "ﾛｰﾀﾘｰｽｲｯﾁ検査"    'Rev0.11DEL
        dfrm.DISP_TITLENAME = "DIP SW 검사"    'Rev0.11ADD
        dfrm.DISP_MESSAGE = mMessage
        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        Call dfrm.frmShow()

        Try
            For i = 1 To 8
                Do
                    If i <= 4 Then
                        If CInt(Mid(DIPSRead, 2, 1)) = 2 ^ (i - 1) Then Exit Do
                    Else
                        If CInt(Mid(DIPSRead, 1, 1)) = 2 ^ (i - 5) Then Exit Do
                    End If
                    'wait push button -------------------
                    If dfrm.DIALOG_RESULT = "NG" Then
                        dfrm.Close()
                        dfrm = Nothing
                        mErrorComment = "ﾃﾞｨｯﾌﾟｽｲｯﾁ検査 異常 "
                        Throw New ApplicationException(mErrorComment)
                    End If
                Loop
                mMessage &= i.ToString
                dfrm.Close()
                dfrm = New frm500_MessagePopUp
                dfrm.DISP_MESSAGE = mMessage
                dfrm.frmShow()
            Next
        Catch
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "ﾃﾞｨｯﾌﾟｽｲｯﾁ検査OK")

    End Sub

    Public Sub I15(ByVal iName As String)
        '==============================================================================================================
        ' Function : I15:MAC ID READ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim mReadMACID As String
        Dim mWriteMACID As String = ""

        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "MacAddress,SerialNo"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""
        Dim ServerPath As String = ServerPathRead()

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            mWriteMACID = sData(0)
            'mLineSerial = sData(1)
            Call WriteLog(False, "書き込みMACｱﾄﾞﾚｽ：" & mWriteMACID)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'ｳｨﾝﾄﾞｳｱﾄﾞﾚｽの変更 $4000が$200で読めるようにする
        cMcm.IO_WRITE(1, 2, "00000211", "0000")
        cMcm.IO_WRITE(1, 2, "00000212", "1010")

        Call cMcm.IO_READ(0, 2, "00000200") : mReadMACID = Strings.Right(cMcm.RecvText, 4)
        Call cMcm.IO_READ(0, 2, "00000202") : mReadMACID &= Strings.Right(cMcm.RecvText, 4)
        Call cMcm.IO_READ(0, 2, "00000204") : mReadMACID &= Strings.Right(cMcm.RecvText, 4)
        Call WriteLog(False, "読み込みMACｱﾄﾞﾚｽ：" & mReadMACID)

        'ｳｨﾝﾄﾞｳｱﾄﾞﾚｽを戻す
        cMcm.IO_WRITE(1, 2, "00000211", "0080")
        cMcm.IO_WRITE(1, 2, "00000212", "0010")

        cMcm = Nothing

        If mReadMACID <> mWriteMACID Then
            mErrorComment = "MACｱﾄﾞﾚｽ読み出し確認異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Else
            Call WriteLog(False, "MACｱﾄﾞﾚｽ読み出し確認OK")
        End If
        
    End Sub

    Public Sub I16(ByVal iName As String)
        '==============================================================================================================
        ' Function : I16:ﾛｰﾀﾘｰｽｲｯﾁ通信設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'メッセージ表示
        dfrm = New frm500_MessagePopUp
        dfrm.DISP_TITLENAME = "ﾛｰﾀﾘｰｽｲｯﾁ設定"
        dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3LExx\RSW2.ppt"
        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        Call dfrm.frmShow()

        Try

            Do
                If RTSRead() = "C0A80002" Then
                    Dim startime As Double = System.Environment.TickCount
                    Do
                        If (System.Environment.TickCount - startime) / 1000 > 1 Then
                            Exit Do '1秒間同じ値であること
                        End If
                    Loop Until RTSRead() <> "C0A80002"

                    If RTSRead() = "C0A80002" Then
                        '1秒間同じ時
                        Exit Do
                    End If

                End If

                DoEvents()
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    dfrm.Close()
                    dfrm = Nothing
                    mErrorComment = "ﾛｰﾀﾘｰｽｲｯﾁ設定 異常 "
                    Throw New ApplicationException(mErrorComment)
                End If
            Loop

        Catch
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "ﾛｰﾀﾘｰｽｲｯﾁ設定OK")

    End Sub

    Public Sub I17(ByVal iName As String)
        '==============================================================================================================
        ' Function : I17:出荷設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'メッセージ表示
        dfrm = New frm500_MessagePopUp
        dfrm.DISP_TITLENAME = "出荷設定"
        dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3LExx\RSW3.ppt"
        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        Call dfrm.frmShow()

        Try
            Do
                DoEvents()
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    dfrm.Close()
                    dfrm = Nothing
                    mErrorComment = "出荷設定 異常 "
                    Throw New ApplicationException(mErrorComment)
                ElseIf dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If
            Loop
        Catch
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
            dfrm.Close()
            dfrm = Nothing
        End Try

        '電源再起動
        'Call I91(iName)
        Call I90("ｽﾛｯﾄ5")

        Try
            If RTSRead() <> "00000000" Then
                mErrorComment = "ﾛｰﾀﾘｰｽｲｯﾁ設定異常:" & RTSRead()
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, "ﾛｰﾀﾘｰｽｲｯﾁ設定OK:" & RTSRead())
            End If
            If DIPSRead() <> "0000" Then
                mErrorComment = "DIPｽｲｯﾁ設定異常:" & DIPSRead()
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, "DIPｽｲｯﾁ設定OK:" & DIPSRead())
            End If
        Catch
            If mErrorComment = "" Then mErrorComment = "出荷設定異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try
        
        Call WriteLog(False, "出荷設定OK")

    End Sub


    Public Sub I21(ByVal iName As String)
        '==============================================================================================================
        ' Function : I21:ﾌｨﾙﾀ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  : 被検査ﾓｼﾞｭｰﾙｽﾛｯﾄ5にて
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Dim cEI As cls130_EI_GAZO
        Dim Model As String = ModelRead()
        Dim mFuncName As String = ""
        Dim mResult As String = ""
        Dim mResult1 As String = ""
        Dim mResult2 As String = ""

        'ﾌｨﾙﾀ画像検査
        Try
            cEI = New cls130_EI_GAZO
            cEI.IPAdd = "192.168.1.30" : cEI.MasterFile = Strings.Left(Model, 9) & "_nonLED.jpg"
            If cEI.MasterFileRead = False Then
                mErrorComment = "ﾏｽﾀﾌｧｲﾙ読み込み異常"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "LEDﾌｨﾙﾀﾁｪｯｸ")
            Call System.Threading.Thread.Sleep(500)
            Call cEI.GAZO_TEST()
            Call WriteLog(False, cEI.ResultData)
            If cEI.RESULT <> "0" Then
                'メッセージ表示
                dfrm = New frm500_MessagePopUp
                'dfrm.DISP_TITLENAME = "LEDﾌｨﾙﾀ検査"                 'Rev0.11DEL
                'dfrm.DISP_MESSAGE = "LEDﾌｨﾙﾀｼﾙｸ印刷は合っていますか？" & vbCrLf & "MODEL=" & Mid(Model, 1, 9)     'Rev0.11DEL
                dfrm.DISP_TITLENAME = "LED필터 검사"                 'Rev0.11ADD
                'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까?" & vbCrLf & "MODEL=" & Mid(Model, 1, 9)     'Rev0.11ADD     Ver 0.23 JDI DEL
                'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까?" & vbCrLf & "MODEL=" & Mid(Model, 3, 7)                     'Ver 0.23 JDI ADD /  Ver 0.31 GYS DEL
                dfrm.DISP_MESSAGE = "필터 형명 확인" & vbCrLf & " = " & Mid(Model, 3, 7)                     'Ver 0.31 GYS ADD
                dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                Call dfrm.frmShow()

                Do
                    'wait push button -------------------
                    If dfrm.DIALOG_RESULT = "NG" Then
                        dfrm.Close()
                        dfrm = Nothing
                        mErrorComment = "ﾌｨﾙﾀ検査 異常 "
                        Throw New ApplicationException(mErrorComment)
                    End If

                    If dfrm.DIALOG_RESULT = "OK" Then
                        Exit Do
                    End If

                    DoEvents()

                Loop

                dfrm.Close()
                dfrm = Nothing

            End If

            Call WriteLog(False, "ﾌｨﾙﾀﾁｪｯｸOK")

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "ﾌｨﾙﾀ検査異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cEI = Nothing
        End Try

        For i = 0 To 1
            Try
                'LANｱﾀﾞﾌﾟﾀdisable
                cPCSP = New cls060_PCNetSPEED
                Call cPCSP.DisableNetAdapter()

                cEI = New cls130_EI_GAZO
                cEI.IPAdd = "192.168.1.30" : cEI.MasterFile = "F3LExx-xx_LED.jpg"
                If cEI.MasterFileRead = False Then
                    mErrorComment = "ﾏｽﾀﾌｧｲﾙ読み込み異常"
                    Throw New ApplicationException(mErrorComment)
                End If
                Call WriteLog(False, "ERR LEDﾁｪｯｸ")
                Call System.Threading.Thread.Sleep(500)
                Call cEI.GAZO_TEST()
                Call WriteLog(False, cEI.RESULT)
                Select Case i
                    Case 0
                        'mFuncName = "RDY点灯,ERR点灯確認"    'Rev0.11DEL
                        mFuncName = "RDY점등,ERR점등확인"    'Rev0.11ADD
                        mResult = "0,1,0,0"
                        mResult1 = "0" : mResult2 = "0"
                    Case 1
                        'mFuncName = "RDY点灯,ERR消灯確認"    'Rev0.11DEL
                        mFuncName = "RDY점등,ERR소등확인"    'Rev0.11ADD
                        mResult = "0,1,0,1"
                        mResult1 = "0" : mResult2 = "1"
                End Select

                'If Strings.Left(cEI.RESULT, 1) <> mResult1 Or Strings.Right(cEI.RESULT, 1) <> mResult2 Then
                If cEI.RESULT <> mResult Then
                    'メッセージ表示
                    dfrm = New frm500_MessagePopUp
                    dfrm.DISP_TITLENAME = mFuncName
                    'dfrm.DISP_MESSAGE = mFuncName & "してください。"   'Rev0.11DEL
                    dfrm.DISP_MESSAGE = mFuncName & "해 주세요."   'Rev0.11ADD
                    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                    Call dfrm.frmShow()

                    Do
                        'wait push button -------------------
                        If dfrm.DIALOG_RESULT = "NG" Then
                            dfrm.Close()
                            dfrm = Nothing
                            mErrorComment = mFuncName & "異常 "
                            Throw New ApplicationException(mErrorComment)
                        End If

                        If dfrm.DIALOG_RESULT = "OK" Then
                            Exit Do
                        End If

                        DoEvents()

                    Loop

                    dfrm.Close()
                    dfrm = Nothing

                End If
checkok:
                Call WriteLog(False, mFuncName & "OK")

            Catch ex As Exception
                If mErrorComment = "" Then mErrorComment = mFuncName & "異常"
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                cEI = Nothing
                Call cPCSP.EnableNetAdapter()
                cPCSP = Nothing
            End Try

            'ERR LEDｺﾝﾄﾛｰﾙ
            Try
                cMcm = New cls120_SP_M_Command
                cMcm.IPAdd = IP_U2
                cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
                cMcm.SlotNo = 1                         'CPU Slot番号
                cMcm.UnitNo = 0                         'ユニット番号

                Select Case i
                    Case 0
                        mFuncName = "ERR LED消灯"
                        cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0E00")
                        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0E00をﾗｲﾄ")
                    Case 1
                        mFuncName = "ERR LED点灯"
                        cMcm.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0D00")
                        Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestComAddress.ToString("X4") & "：$0D00をﾗｲﾄ")
                End Select

                cMcm.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")
                Call WriteLog(False, "ｱﾄﾞﾚｽ$" & mTestExeAddress.ToString("X4") & "：$8200をﾗｲﾄ")
                If FlagCheck() = True Then
                    Call WriteLog(False, mFuncName & "OK")
                Else
                    mErrorComment = mFuncName & "異常"
                    Throw New ApplicationException(mErrorComment)
                End If
            Catch
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                cMcm = Nothing
            End Try

        Next

        'RDY,LNK,ERR LED点灯目視確認
        Try
            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3LExx\RDYLNKERR_LED.ppt"
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

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

    End Sub


    Public Sub I30(ByVal iName As String)
        '==============================================================================================================
        ' Function : I30:Ethernet通信検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "Ethernet通信検査中・・・")

        Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
        Dim adapter As NetworkInterface

        Dim sKEY As String
        Dim NetAdpMAC As String
        Dim Ret As Boolean
        Dim Ret1 As Boolean
        Dim RegStr As String

        Dim Model As String = ModelRead()
        Dim RecvDataLE As String = ""

        'ﾈｯﾄﾜｰｸｱﾀﾞﾌﾟﾀ確認 --------------------------------------------------------------------
        For Each adapter In adapters
            NetAdpMAC = adapter.GetPhysicalAddress.ToString
        Next adapter

        '設定 --------------------------------------------------------------------
        'LAN Speed Control
        sKEY = "*SpeedDuplex"

        '速度10M Fullに変更 ------------------------------------------------------
        Call WriteLog(False, "通信速度10M Fullに変更中・・・")

        DoEvents()

        RegStr = "2" '10M Full
        cPCSP = New cls060_PCNetSPEED

        Ret1 = cPCSP.VBSetRegistryString(RegStr, sKEY)
        Ret = cPCSP.DisEnableNetAdapter

        cPCSP = Nothing
        Call WriteLog(False, "速度変更　反映待機中・・・")
        Call Insfrm.WaitStart(7)

        Call WriteLog(False, "通信速度10M Fullでの通信確認中・・・")
        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_LE
            Call cLnk.F3PCLinkCommunication("01INF3,0", RecvDataLE)
            Debug.Print(RecvDataLE)
            Call WriteLog(False, "Send：01INF3,0")
            Call WriteLog(False, "Recv：" & Replace(RecvDataLE, vbCrLf, ""))
            cLnk = Nothing
        Catch
            mErrorComment = "10M通信異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, "送受信10M OK")
        Call WriteLog(False, "通信速度10M Fullでの通信　OK")

        Call WriteLog(False, "")
        Call Insfrm.WaitStart(1)

        '速度100M Fullに変更 -----------------------------------------------------
        Call WriteLog(False, "通信速度100M Fullに変更中・・・")

        RegStr = "4" '100M Full

        cPCSP = New cls060_PCNetSPEED

        Ret1 = cPCSP.VBSetRegistryString(RegStr, sKEY)
        Ret = cPCSP.DisEnableNetAdapter

        cPCSP = Nothing
        Call WriteLog(False, "速度変更　反映待機中・・・")
        Call Insfrm.WaitStart(7)

        If Mid(Model, 3, 4) = "LE01" Then Exit Sub 'LE01のときは10Mのみ検査

        Call WriteLog(False, "通信速度100M Fullでの通信確認中・・・")
        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_LE
            Call cLnk.F3PCLinkCommunication("01INF3,0", RecvDataLE)
            Debug.Print(RecvDataLE)
            Call WriteLog(False, "Send：01INF3,0")
            Call WriteLog(False, "Recv：" & Replace(RecvDataLE, vbCrLf, ""))
            cLnk = Nothing
        Catch
            mErrorComment = "100M通信異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, "通信速度100M Fullでの通信　OK")

        '正常終了 ----------------------------------------------------------------
        Call WriteLog(False, "Ethernet通信検査 完了")
        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I50(ByVal iName As String)
        '==============================================================================================================
        ' Function : I50:SLOT/XDFAIL TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I50(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I52(ByVal iName As String)
        '==============================================================================================================
        ' Function : I52:BCLR TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I52(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I53(ByVal iName As String)
        '==============================================================================================================
        ' Function : I53:IOﾊﾞｽ-XBRﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I53(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

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
        'ﾒｲﾝCPU（SP7x）のXIRQをﾏｽｸしてﾓｼﾞｭｰﾙ側からXIRQ0～3を上げ下げして
        'IOﾊﾞｽﾚｼﾞｽﾀにてXIRQの変化がﾓﾆﾀできること

        Dim mReadAddress As String = ""
        Dim mRegSlotAddress As String = "00000200"
        Dim mRegIRQMon As String = "00000207"
        Dim mRegBRMon As String = "00000230"
        Dim mRegMODFail As String = "00000224"
        Dim mRegLFail As String = "00000217"
        Dim mIRQ As Integer

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'ﾒｲﾝCPUのXIRQﾏｽｸ
        Call WriteLog(False, "ﾒｲﾝCPUのXIRQﾏｽｸ")
        Call cMcm.MEM_WRITE(1, "3880021E", "00")

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Try
            For i As Integer = 3 To 0 Step -1
                mIRQ = CInt(2 ^ i)
                cMcm.IO_WRITE(1, 2, "00000208", mIRQ.ToString & "080")

                Call WriteLog(False, "XIRQ" & CStr(i) & "確認")
                Call WriteLog(False, "Send : " & "ｱﾄﾞﾚｽ$207ﾘｰﾄﾞ（0x07のｽﾛｯﾄ5IRQﾚｼﾞｽﾀ）")
                Call cMcm.IO_READ(1, 1, mRegIRQMon)
                Call WriteLog(False, "Recv : " & cMcm.RecvText)

                Select Case i
                    Case 0 : mResStr = "1"
                    Case 1 : mResStr = "2"
                    Case 2 : mResStr = "4"
                    Case 3 : mResStr = "8"
                End Select

                If Mid(Strings.Right(cMcm.RecvText, 4), 1, 1) <> mResStr Then
                    mErrorComment = "XIRQ異常" & i.ToString
                    Throw New ApplicationException(mErrorComment)
                End If

                For j = 0 To 3
                    Select Case j
                        Case 0 : mReadAddress = mRegBRMon
                        Case 1 : mReadAddress = mRegSlotAddress
                        Case 2 : mReadAddress = mRegMODFail
                        Case 3 : mReadAddress = mRegLFail
                    End Select

                    If j = 3 Then
                        Call cMcm.IO_READ(0, 1, mReadAddress)
                    Else
                        Call cMcm.IO_READ(1, 1, mReadAddress)
                    End If

                    Call WriteLog(False, "Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                    Call WriteLog(False, "Recv : " & cMcm.RecvText)

                    If j = 1 Then
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
                            mErrorComment = "XIRQとその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    ElseIf j = 0 Then
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "80" Then
                            mErrorComment = "XIRQとその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Else
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                            mErrorComment = "XIRQとその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    End If

                Next j

                Call WriteLog(False, "XIRQ" & i.ToString & " OK")

            Next
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XIRQ異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm.IO_WRITE(1, 2, "00000208", "0080")
            'ﾒｲﾝCPUのXIRQﾏｽｸ解除
            Call WriteLog(False, "ﾒｲﾝCPUのXIRQﾏｽｸ解除")
            Call cMcm.MEM_WRITE(1, "3880021E", "88")
            cBUS = Nothing
            cMcm = Nothing
        End Try

        Call WriteLog(False, "XIRQ：OK")


    End Sub

    Public Sub I56(ByVal iName As String)
        '==============================================================================================================
        ' Function : I56:IOﾊﾞｽ-XLFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I56(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try
    End Sub

    Public Sub I57(ByVal iName As String)
        '==============================================================================================================
        ' Function : I57:IOﾊﾞｽ-XPFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I57(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I58(ByVal iName As String)
        '==============================================================================================================
        ' Function : I58:IOﾊﾞｽ-XSFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I58(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I59(ByVal iname As String)
        '==============================================================================================================
        ' Function : I59:IOﾊﾞｽ検査前準備
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I59(iname)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
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

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I90(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I91(ByVal iName As String)
        '==============================================================================================================
        ' Function : I91:ﾕﾆｯﾄ電源OFF
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I91(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

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
    Public Sub IA1(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA1:Mac Address Port
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Dim cSP7x As cls2D2_F3SP7x_Inspection
        cSP7x = New cls2D2_F3SP7x_Inspection
        Try
            Call cSP7x.IA1()
        Catch ex As Exception
            mErrorComment = cSP7x.ERR_COMMENT : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP7x = Nothing
        End Try

    End Sub

    Public Sub IA2(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA2:MACｱﾄﾞﾚｽ重複確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim cSP7x As cls2D2_F3SP7x_Inspection
        cSP7x = New cls2D2_F3SP7x_Inspection
        Try
            Call cSP7x.IA2(iName)
        Catch ex As Exception
            mErrorComment = cSP7x.ERR_COMMENT : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP7x = Nothing
        End Try

    End Sub


    Public Sub IP1(ByVal iName As String)
        '==============================================================================================================
        ' Function : IP1:PrintLabel
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim cRPxx As cls2D3_F3RPxx_Inspection

        cRPxx = New cls2D3_F3RPxx_Inspection
        Try
            Call cRPxx.IP1(iName)           'RPxxと同じ主銘板ﾌｫｰﾏｯﾄ
        Catch ex As Exception
            mErrorComment = cRPxx.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cRPxx = Nothing
        End Try

    End Sub

    Private Sub RePrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : 主銘板再印字 
        'Remarks  :
        '==============================================================================================================

        Call IP1("主銘板再印字")

    End Sub

    Private Function FlagCheck() As Boolean

        Dim cMcm1 As cls120_SP_M_Command
        Dim t, t1 As Integer
        Dim TimeOut As Integer
        Dim TotalWait As Double

        cMcm1 = New cls120_SP_M_Command
        cMcm1.IPAdd = IP_U2                      'CPU用LE経由
        cMcm1.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm1.SlotNo = 1                         'CPU Slot番号
        cMcm1.UnitNo = 0                         'ユニット番号

        If mTimeOut = 0 Then
            TimeOut = 5
        Else
            TimeOut = mTimeOut
        End If

        t = System.Environment.TickCount
        Do
            t1 = System.Environment.TickCount
            TotalWait = CInt((t1 - t) / 1000)
            If TotalWait >= TimeOut Then
                mErrorComment = "timeout" : Exit Function
            End If
            Call cMcm1.IO_READ(0, 2, "00000234")
            DoEvents()
            Select Case Strings.Right(cMcm1.RecvText, 4)
                Case "0000" : Exit Do
                Case "0074", "0075", "0076", "0077", "0078", "0079", "007A", "007B" : Exit Function
                Case Else
            End Select
        Loop

        FlagCheck = True
        cMcm1 = Nothing

    End Function

    Private Function RTSRead() As String
        'ﾛｰﾀﾘｰｽｲｯﾁの状態読み込み
        RTSRead = ""
        Dim cMcm2 As cls120_SP_M_Command

        cMcm2 = New cls120_SP_M_Command
        cMcm2.IPAdd = IP_U2
        cMcm2.RemPort = 12290
        cMcm2.SlotNo = 1
        cMcm2.UnitNo = 0

        cMcm2.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0200")
        cMcm2.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")

        If FlagCheck() = True Then
            cMcm2.IO_READ(0, 2, "0000023C") : RTSRead = Strings.Right(cMcm2.RecvText, 4)
            cMcm2.IO_READ(0, 2, "0000023E") : RTSRead &= Strings.Right(cMcm2.RecvText, 4)
        End If

        cMcm2 = Nothing

    End Function

    Private Function DIPSRead() As String
        'ﾃﾞｨｯﾌﾟｽｲｯﾁの状態読み込み
        DIPSRead = ""
        Dim cMcm3 As cls120_SP_M_Command

        cMcm3 = New cls120_SP_M_Command
        cMcm3.IPAdd = IP_U2
        cMcm3.RemPort = 12290
        cMcm3.SlotNo = 1
        cMcm3.UnitNo = 0

        cMcm3.IO_WRITE(0, 2, mTestComAddress.ToString("X8"), "0300")
        cMcm3.IO_WRITE(0, 2, mTestExeAddress.ToString("X8"), "8200")

        If FlagCheck() = True Then
            cMcm3.IO_READ(0, 2, mTestResAddress.ToString("X8")) : DIPSRead = Strings.Right(cMcm3.RecvText, 4)
        End If

        cMcm3 = Nothing

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

    Private Function ServerPathRead() As String

        'serverパスの読込み
        Dim dbName As String = mAppPath & "010 EquipmentData\EquipmentSetting.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "BackupDrivePath"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        ServerPathRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblEquipmentSetting"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            ServerPathRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function

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

End Class
