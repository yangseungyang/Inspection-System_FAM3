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


Public Class cls2D2_F3SP7x_Inspection
    'Class ================================================
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cLnk As cls080_F3LinkToLE
    'Dim cDMM As cls101_ComRS232C
    Dim cISPComm As cls101_ComRS232C
    Dim cWin As cls001_Window_Control
    Dim cMac As cls031_MacAddress
    Dim cSer As cls030_LineSerial
    Dim cLbl As cls011_Excel_Operation
    Dim cMcm As cls120_SP_M_Command
    Dim cPCSP As cls060_PCNetSPEED
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl
    Dim cTEST As cls121_SP_TEST_MODE
    Dim cBUBU As cls123_BUBU_IOBUS_TEST

    Dim dfrm As frm500_MessagePopUp

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String
    Dim mTestSpec As String    'Rev0.20Add
    Dim mTestMeasure As String 'Rev0.20Add

    Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3
    Dim IP_U2 As String = "192.168.1.11"        'B側用LE
    Dim IP_SP As String = "192.168.0.2"         '被検査用SP

    'Dim rs232cDMMPort As String = "COM3"        'DMMのRS232Cﾎﾟｰﾄ
    Dim iName As String
    Dim rcvMsg As String

    Public usc As UDPState

    'Property =====================================================================================
    'Error message --------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public ReadOnly Property TestSpec() As String    'Rev0.20add
        Get
            TestSpec = mTestSpec
        End Get
    End Property

    Public ReadOnly Property TestMeasure() As String   'Rev0.20Add
        Get
            TestMeasure = mTestMeasure
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
            Case "I05" : Call I05(iName)            'I05:TestMode移行
            Case "I06" : Call I06(iName)            'I06:TestMode終了
            Case "I07" : Call I07(iName)            'I07:CPU-ﾒﾓﾘ検査
            Case "I08" : Call I08(iName)            'I08:SPU3-ﾒﾓﾘ検査
            Case "I09" : Call I09(iName)            'I09:SPU3検査
            Case "I10" : Call I10(iName)            'I10:ﾗﾀﾞｰﾌﾗｯｼｭ検査
            Case "I11" : Call I11(iName)            'I11:割込み信号ﾃｽﾄ
            Case "I12" : Call I12(iName)            'I12:ｸﾛｯｸﾃｽﾄ
            Case "I13" : Call I13(iName)            'I13:RTCﾃｽﾄ
            Case "I14" : Call I14(iName)            'I14:電源電圧検査
            Case "I15" : Call I15(iName)            'I15:Ethernet通信ﾃｽﾄ
            Case "I16" : Call I16(iName)            'I16:ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟﾃｽﾄ               
            Case "I17" : Call I17(iName)            'I17:USB通信ﾃｽﾄ
            Case "I18" : Call I18(iName)            'I18:IOﾊﾞｽ-ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄ
            Case "I19" : Call I19(iName)            'I19:IOﾊﾞｽ-XACSﾃｽﾄ
            Case "I20" : Call I20(iName)            'I20:IOﾊﾞｽ-XBRﾃｽﾄ
            Case "I21" : Call I21(iName)            'I21:IOﾊﾞｽ-XDFAILﾃｽﾄ
            Case "I22" : Call I22(iName)            'I22:IOﾊﾞｽ-割込みﾃｽﾄ
            Case "I23" : Call I23(iName)            'I23:IOﾊﾞｽ-XLFAILﾃｽﾄ
            Case "I24" : Call I24(iName)            'I24:IOﾊﾞｽ-XPFAILﾃｽﾄ
            Case "I25" : Call I25(iName)            'I25:IOﾊﾞｽ-XSFAILﾃｽﾄ
            Case "I26" : Call I26(iName)            'I26:LED検査
            Case "I27" : Call I27(iName)            'I27:ﾛｰﾀﾘｰｽｲｯﾁLED検査
            Case "I28" : Call I28(iName)            'I28:SDｶｰﾄﾞ検査-ｶｰﾄﾞ電源/過電流保護
            Case "I29" : Call I29(iName)            'I29:SDｶｰﾄﾞ検査-ﾗｲﾄﾌﾟﾛﾃｸﾄ検査
            Case "I30" : Call I30(iName)            'I30:ﾗﾀﾞｰ検査
            Case "I31" : Call I31(iName)            'I31:EthernetLEDﾃｽﾄ
            Case "I32" : Call I32(iName)            'I32:ﾃﾞｰﾀ書込み
            Case "I33" : Call I33(iName)            'I33:SDｶｰﾄﾞ検査-ﾉﾝﾗｲﾄﾌﾟﾛﾃｸﾄ検査
            Case "I34" : Call I34(iName)            'I34:IOﾊﾞｽ検査前準備
            Case "I35" : Call I35(iName)            'I35:出荷設定
            Case "I36" : Call I36(iName)            'I36:ﾗﾀﾞｰﾌﾗｯｼｭ部出荷設定
            Case "I37" : Call I37(iName)            'I37:RTC確認

            Case "I40" : Call I40(iName)            'I40:電源電圧測定始業点検DMM

            Case "IA2" : Call IA2(iName)            'IA2:MACｱﾄﾞﾚｽ重複確認

            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF

            Case "I92" : Call I92(iName)            'I92:検査開始時処理
            Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save        'Rev 0.26 KSJ ADD
            Case "IP1" : Call IP1(iName)            'IP1:PrintLabel
            Case "IP2" : Call IP2(iName)            'IP2:ﾃｽﾄ印字

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
        Dim wTime As Double = 0
        Dim writeModel As String = ModelRead()
        Dim WriteData As String = ""
        Dim RecvDataLE As String = ""
        Dim t, t1 As Integer
        Dim tWait As Integer

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call I90(iName)                             'ｱﾄﾞｵﾝCPUで電源ON
        Call WaitStart(3)

        If writeModel.Length > 9 Then writeModel = Strings.Left(writeModel, 9)

        'VCC_B側のLE経由でﾏｽﾀCPUにPCﾘﾝｸｺﾏﾝﾄﾞ送信
        For i = 1 To writeModel.Length
            WriteData &= CStr(Hex(Asc(Mid(writeModel, i, 1))))
        Next

        '-------------------- Ver 0.22 JDI ADD Start ---------------------
        If InStr(ModelRead(), "C12") <> 0 Then
            For i = 1 To "/C12".Length
                WriteData &= CStr(Hex(Asc(Mid("/C12", i, 1))))
            Next
        End If
        '--------------------- Ver 0.22 JDI ADD End ----------------------

        WriteData &= "00"           '端数の最後には”00”を書く

        Dim dfrm As New frm501_StatusDisp

        dfrm.Text = "SP7xﾌｧｰﾑ書き込み中。中止できません。"
        Call WriteLog(False, "ﾌｧｰﾑ書き込み開始")
        tWait = 180 * 1000                         '最大3分00秒待ちます 180秒×1000
        t = System.Environment.TickCount

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U2
            Call cLnk.F3PCLinkCommunication("01" & "WWRZ602,01,0000", RecvDataLE)

            '-------------------- Ver 0.22 JDI ADD Start ---------------------
            If InStr(ModelRead(), "C12") <> 0 Then
                Call cLnk.F3PCLinkCommunication("01" & "WWRZ603,07," & WriteData, RecvDataLE)
                Call cLnk.F3PCLinkCommunication("01" & "WRDZ603,07", RecvDataLE)
            Else
                Call cLnk.F3PCLinkCommunication("01" & "WWRZ603,05," & WriteData, RecvDataLE)
                Call cLnk.F3PCLinkCommunication("01" & "WRDZ603,05", RecvDataLE)
            End If
            '--------------------- Ver 0.22 JDI ADD End ----------------------

            'Call cLnk.F3PCLinkCommunication("01" & "WWRZ603,05," & WriteData, RecvDataLE)     'Ver 0.22 JDI DEL
            'Call cLnk.F3PCLinkCommunication("01" & "WRDZ603,05", RecvDataLE)                  'Ver 0.22 JDI DEL

            Call cLnk.F3PCLinkCommunication("01" & "WWRZ601,01,0007", RecvDataLE)

            '書き込みｽﾃｰﾀｽ確認
            Do
                DoEvents()
                t1 = System.Environment.TickCount
                dfrm.TextBox1.Text = "あと約" & CInt((tWait - (t1 - t)) / 1000) & "秒です" & vbCrLf
                If t1 - t > 120000 Then
                    Call cLnk.F3PCLinkCommunication("01" & "WRDZ601,01", RecvDataLE)
                    dfrm.TextBox1.Text &= "Send : " & "01" & "WRDZ601,01" & vbCrLf
                    dfrm.TextBox1.Text &= "Recv : " & RecvDataLE
                    If Mid(RecvDataLE, 1, 8) = "11OK0000" Then Exit Do
                End If
                dfrm.ProgressBar1.Value = CInt((t1 - t) * 100 / tWait)
                Call dfrm.Show()
            Loop
            '書き込み結果確認
            Call cLnk.F3PCLinkCommunication("01" & "WRDZ602,01", RecvDataLE)
            If Mid(RecvDataLE, 1, 8) <> "11OK0000" Then
                mErrorComment = "ﾌｧｰﾑ書き込みｴﾗｰ ｺｰﾄﾞ:" & Strings.Right(RecvDataLE, 4)
                mTestSpec = "11OK0000"      'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            MsgBox("エラー")
            If mErrorComment = "" Then mErrorComment = "ﾌｧｰﾑ書き込みｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
            dfrm.Dispose()
            dfrm = Nothing
        End Try

        Call WaitTimer(3)
        Call WriteLog(False, "ﾌｧｰﾑ書き込み正常終了")

        Call I91(iName)         'ﾓｼﾞｭｰﾙ電源OFF

        Call I90("1")           '被検査ﾓｼﾞｭｰﾙをﾒｲﾝCPUで電源ON

        Call WaitStart(5)

    End Sub

    Public Sub I03(ByVal iName As String)
        '==============================================================================================================
        ' Function : I03:ﾓｼﾞｭｰﾙIDﾁｪｯｸ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendCom As String
        Dim Model As String = ModelRead()
        Dim ResStr As String = ""

        If InStr(iName, "I03") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================
        SendCom = "TS,13,0,0,"
        If Strings.Left(Model, 9) = "F3SP76-7N" Then RESSTR = "F3SP76-7N"
        If Strings.Left(Model, 9) = "F3SP71-4N" Then RESSTR = "F3SP71-4N"
        If Strings.Left(Model, 9) = "F3SP76-7S" Then ResStr = "F3SP76-7S"
        If Strings.Left(Model, 9) = "F3SP71-4S" Then ResStr = "F3SP71-4S"
        'If Strings.Left(Model, 9) = "F3SPV9-7S" Then ResStr = "F3SP76-7S" 'Rev0.12ADD Rev0.15DEL
        If Strings.Left(Model, 9) = "F3SPV9-7S" Then ResStr = "F3SPV9-7S" 'Rev0.15ADD

        If TSSendRecv(SendCom, 5) <> "TS,13,0,0,0," & ResStr & "," Then
            'mErrorComment = "ﾓｼﾞｭｰﾙID確認ｴﾗｰ"           'Rev0.20del
            mErrorComment = "ﾓｼﾞｭｰﾙID確認ｴﾗｰ" & SendCom 'Rev0.20add
            MsgBox(mErrorComment)
            Call WriteLog(False, mErrorComment)
            mTestSpec = ResStr     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "ﾓｼﾞｭｰﾙID確認　OK")

    End Sub

    Public Sub I04(ByVal iName As String)
        '==============================================================================================================
        ' Function : I04:ROMSUMﾁｪｯｸ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String
        Dim RESSTR As String = ""
        Dim Model As String = ModelRead()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Dim dbName As String
        Dim rDTkey As String = ""
        Dim tbName As String = ""
        Dim rField As String = ""
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim ROM_SUM76 As String = ""
        Dim ROM_SUM71 As String = ""
        Dim ROM_SUMBT As String = ""

        'v0.07 SP7x-xS対応
        Select Case Mid(Model, 9, 1)
            Case "N"
                tbName = "tblCheckSUMN"
                rField = "CheckSUM76,CheckSUM71,CheckSUMBT"                                  'Ver 0.22 JDI ADD
            Case "S"
                tbName = "tblCheckSUMS"
                rField = "CheckSUM76,CheckSUM71,CheckSUMBT,CheckSUM76C12,CheckSUM71C12"       'Ver 0.22 JDI ADD
                rField = rField & ",CheckSUM76CS,CheckSUM71CS,CheckSUMBTCS"                 'Ver 0.49 HSJ ADD (/CS用CS列を追加)
        End Select

        'チェックサムの読込み
        dbName = mAppPath & "520 Product Firmware\SP7X\FirmName.mdb"
        'tbName = "tblCheckSUM"
        'rField = "CheckSUM76,CheckSUM71,CheckSUMBT"                                 'Ver 0.22 JDI DEL

        '-----Rev0.12ADD-----
        If Mid(Model, 1, 9) = "F3SPV9-7S" Then
            tbName = "tblCheckSUMVP"
            rField = "CheckSUMV9,CheckSUMxx,CheckSUMBT"
        End If
        '---------------------

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)
            ROM_SUM76 = sData(0) : ROM_SUM71 = sData(1) : ROM_SUMBT = sData(2)
            '------------------- Ver 0.22 JDI ADD Start ---------------------
            If InStr(Model, "C12") <> 0 Then
                ROM_SUM76 = sData(3) : ROM_SUM71 = sData(4)
            End If
            '-------------------- Ver 0.22 JDI ADD End ----------------------
            If InStr(Model, "/CS") <> 0 Then ROM_SUM76 = sData(5) : ROM_SUM71 = sData(6) : ROM_SUMBT = sData(7) 'Ver 0.49 HSJ ADD

        Catch ex As Exception
            cADO = Nothing
        End Try

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "CPU周辺-ROM SUM確認中・・・")

        'CPU周辺-ROM SUM(Boot)確認 -------------------------
        Call WriteLog(False, "CPU周辺-ROM SUM(BOOT)確認中・・・約２秒待ち")

        Try
            rcv = TSSendRecv("TS,29,0,0,", 10)
            If rcv <> "TS,29,0,0,0," & ROM_SUMBT & "," Then
                mErrorComment = "CPU周辺-ROM SUM(BOOT)確認　異常 " & rcv
                Call WriteLog(False, mErrorComment)
                mTestSpec = "TS,29,0,0,0," & ROM_SUMBT & ","    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "CPU周辺-ROM SUM(BOOT)確認 OK ")

        'CPU周辺-ROM SUM(Firm)確認 -------------------------
        WriteLog(False, "CPU周辺-ROM SUM(Firm)確認中・・・約２秒待ち")

        Try
            rcv = TSSendRecv("TS,4,0,0,0,0,", 10)
            'If InStr(Model, "F3SP76") <> 0 Then         'Mid(Model, 1, 6) = "F3SP66" Then    Rev0.12DEL
            If InStr(Model, "F3SP76") <> 0 Or InStr(Model, "F3SPV9") <> 0 Then            'Rev0.12ADD
                'ResultData.ROM_SUM66 = Mid(rcv, 12, Len(rcv) - 12)
                If rcv <> "TS,4,0,0,0," & ROM_SUM76 & "," Then
                    mErrorComment = "CPU周辺-ROM SUM(Firm)確認　異常 " & rcv
                    Call WriteLog(False, mErrorComment)
                    mTestSpec = "TS,4,0,0,0," & ROM_SUM76 & ","    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                Call WriteLog(False, "CPU周辺-ROM SUM(Firm)確認 OK ")
            Else
                'ResultData.ROM_SUM67 = Mid(rcv, 12, Len(rcv) - 12)
                If rcv <> "TS,4,0,0,0," & ROM_SUM71 & "," Then
                    mErrorComment = "CPU周辺-ROM SUM(Firm)確認　異常 " & rcv
                    Call WriteLog(False, mErrorComment)
                    mTestSpec = "TS,4,0,0,0," & ROM_SUM71 & ","   'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                Call WriteLog(False, "CPU周辺-ROM SUM(Firm)確認 OK ")
            End If

        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

    End Sub

    Public Sub I05(ByVal iName As String)
        '==============================================================================================================
        ' Function : I05:TestMode移行
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Dim SendCom As String = ""
        Dim RESSTR As String = ""
        Dim i As Integer = 0
        Dim j As Integer = 0

        If InStr(iName, "I05") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================
retry:
        If TSSendRecv("SP TEST MODE START", 3) <> "TS,FF,0,0,F2," Then
            If i > 5 Then
                Call I91(iName)     'ﾓｼﾞｭｰﾙ電源OFF
                Call I90("1")       'ﾓｼﾞｭｰﾙ電源ON
                j = j + 1 : i = 0
                If j > 2 Then
                    mErrorComment = "ﾃｽﾄﾓｰﾄﾞ移行ﾘﾄﾗｲﾀｲﾑｱｳﾄｴﾗｰ"
                    'MsgBox(mErrorComment)
                    Call WriteLog(False, mErrorComment)
                    mTestSpec = "TS,FF,0,0,F2,"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                GoTo retry
            End If
            i = i + 1
            GoTo retry
            'mErrorComment = "ﾃｽﾄﾓｰﾄﾞ移行NG"
            'MsgBox(mErrorComment)
            'Call WriteLog(False, mErrorComment)
            'Throw New ApplicationException(mErrorComment)
        Else

            Call WaitStart(3)

            SendCom = "TS,1,0,0,"

            'If TSSendRecv(SendCom, 5) <> "TS,1,0,0,0,0,0,0,0," Then         'Ver 0.38 LCH DEL
            If TSSendRecv(SendCom, 10) <> "TS,1,0,0,0,0,0,0,0," Then         'Ver 0.38 LCH ADD
                If i > 3 Then
                    mErrorComment = "ﾃｽﾄﾓｰﾄﾞ移行ﾘﾄﾗｲﾀｲﾑｱｳﾄｴﾗｰ"
                    MsgBox(mErrorComment)
                    Call WriteLog(False, mErrorComment)
                    mTestSpec = "TS,1,0,0,0,0,0,0,0,"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                i = i + 1
                GoTo retry
            End If

            Call WriteLog(False, "TESTMODE移行OK")

        End If

    End Sub

    Public Sub I06(ByVal iName As String)
        '==============================================================================================================
        ' Function : I06:TestMode終了
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String

        If InStr(iName, "I06") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================

        'Test
        Call WriteLog(False, "TestMode終了中・・・")

        Try
            rcv = TSSendRecv("TS,2,0,0,", 5)
            If rcv <> "TS,2,0,0,0," Then
                mErrorComment = "TestMode終了終了　異常" & rcv    'Rev0.20add
                mTestSpec = "TS,2,0,0,0,"    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            'mErrorComment = "TestMode終了終了　異常"   'Rev0.20del
            If mErrorComment = "" Then mErrorComment = "TestMode終了終了　異常" 'Rev0.20add
            Debug.Print(mErrorComment)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally

        End Try

        Call WaitStart(5)

        Call WriteLog(False, "TestMode終了OK")

    End Sub

    Public Sub I07(ByVal iName As String)
        '==============================================================================================================
        ' Function : I07:CPU-ﾒﾓﾘ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        'CPUﾜｰｸSDRAMﾃｽﾄ,CPU内蔵RAMﾃｽﾄ       ﾃｽﾄｺﾏﾝﾄﾞ未定！

        'For i = 0 To 1
        '    Select Case i
        '        Case 0 : FuncName = "CPUﾜｰｸSDRAMﾃｽﾄ" : SendCom = "TS,3,0,0,0,後日設定" : RecvCom = ""
        '        Case 1 : FuncName = "CPU内蔵RAMﾃｽﾄ" : SendCom = "TS,3,0,0,0,FFF80000,FFF87FFF," : RecvCom = ""
        '    End Select
        '    Call WriteLog(False, FuncName & "中・・・約２秒待ち")

        '    Try
        '        rcv = TSSendRecv(SendCom, 5)
        '        If rcv <> RecvCom Then
        '            mErrorComment = FuncName & "　異常 " & rcv
        '            Throw New ApplicationException(mErrorComment)
        '        End If
        '    Catch ex As Exception
        '        If mErrorComment = "" Then mErrorComment = FuncName & "異常"
        '        Call WriteLog(False, mErrorComment)
        '        Throw New ApplicationException(mErrorComment)
        '    End Try
        '    Call WriteLog(False, FuncName & " OK ")

        'Next i

        For i = 0 To 1
            Select Case i
                Case 0 : FuncName = "CPUﾜｰｸSDRAMﾃｽﾄCPU内臓RAMﾃｽﾄ1" : SendCom = "TS,2A,0,0," : RecvCom = "TS,2A,0,0,0,"
                Case 1 : FuncName = "CPUﾜｰｸSDRAMﾃｽﾄCPU内臓RAMﾃｽﾄ2" : SendCom = "TS,2B,0,0," : RecvCom = "TS,2B,0,0,0,0,0,0,"
            End Select
            Call WriteLog(False, FuncName & "中・・・")

            Try
                rcv = TSSendRecv(SendCom, 10)
                If rcv <> RecvCom Then
                    mErrorComment = FuncName & "　異常 " & rcv
                    mTestSpec = RecvCom                           'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
            Catch ex As Exception
                If mErrorComment = "" Then mErrorComment = FuncName & "異常"
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLog(False, FuncName & " OK ")

            If i = 0 Then Call WaitStart(10)

        Next i

    End Sub

    Public Sub I08(ByVal iName As String)
        '==============================================================================================================
        ' Function : I08:SPU3-ﾒﾓﾘ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Dim rcv As String
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        'SPU3ﾃﾞﾊﾞｲｽSDRAMﾃｽﾄ,SPU3ﾌﾟﾛｸﾞﾗﾑSDRAMﾃｽﾄ,SPU3ﾊﾞｯｸｱｯﾌﾟSRAMﾃｽﾄ
        For i = 0 To 2
            Select Case i
                Case 0 : FuncName = "SPU3ﾃﾞﾊﾞｲｽSDRAMﾃｽﾄ" : SendCom = "TS,3,0,0,0,34000000,347FFFFF," : RecvCom = "TS,3,0,0,0,0,347FFFFC,0,0,"
                Case 1 : FuncName = "SPU3ﾌﾟﾛｸﾞﾗﾑSDRAMﾃｽﾄRAMﾏｰﾁﾝｸﾞﾃｽﾄ" : SendCom = "TS,3,0,0,0,36000000,36FFFFFF," : RecvCom = "TS,3,0,0,0,0,36FFFFFC,0,0,"
                    'Case 2 : FuncName = "SPU3ﾊﾞｯｸｱｯﾌﾟSRAMﾃｽﾄ準備:KEYｵｰﾌﾟﾝ" : SendCom = "TS,7,0,0,39800F00,4,53505533," : RecvCom = "TS,7,0,0,1,4,39800F00,0,"
                Case 2
                    FuncName = "SPU3ﾊﾞｯｸｱｯﾌﾟSRAMﾃｽﾄRAMﾏｰﾁﾝｸﾞﾃｽﾄ"
                    Select Case Strings.Left(Model, 6)
                        Case "F3SP76" : SendCom = "TS,3,0,0,0,3B400000,3B5FFFFF," : RecvCom = "TS,3,0,0,0,0,3B5FFFFC,0,0,"
                        Case "F3SP71" : SendCom = "TS,3,0,0,0,3B400000,3B4FFFFF," : RecvCom = "TS,3,0,0,0,0,3B4FFFFC,0,0,"
                        Case "F3SPV9" : SendCom = "TS,3,0,0,0,3B400000,3B5FFFFF," : RecvCom = "TS,3,0,0,0,0,3B5FFFFC,0,0,"    'Rev0.12ADD
                        Case Else : mErrorComment = "形名登録なし" : Call WriteLog(False, mErrorComment) : Throw New ApplicationException(mErrorComment)
                    End Select
            End Select

            Call WriteLog(False, FuncName & "中・・・")

            Try
                rcv = TSSendRecv(SendCom, 30)
                If rcv <> RecvCom Then
                    rcv = TSSendRecv(SendCom, 30)
                    If rcv <> RecvCom Then
                        mErrorComment = FuncName & "　異常 " & rcv
                        mTestSpec = RecvCom    'Rev0.20add
                        Throw New ApplicationException(mErrorComment)
                    End If
                End If
            Catch ex As Exception
                If mErrorComment = "" Then mErrorComment = FuncName & "異常"
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLog(False, FuncName & " OK ")


        Next i

    End Sub

    Public Sub I09(ByVal iName As String)
        '==============================================================================================================
        ' Function : I09:SPU3検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String
        Dim MeasureValue As Integer
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        'SPUｸﾛｯｸﾃｽﾄ,SPUﾃｽﾄ,SPU3ｷｬｯｼｭﾃｽﾄ,SPU3内部RAMﾃｽﾄ      
        For i = 0 To 4
            Select Case i
                Case 0 : FuncName = "SPUｸﾛｯｸﾃｽﾄ" : SendCom = "TS,5,0,0,1,AAA0,AE60,0," : RecvCom = "TS,5,0,0,0,1,0,"
                Case 1 : FuncName = "SPU自己診断ﾃｽﾄ" : SendCom = "TS,5,0,0,0,0,0,0," : RecvCom = "TS,5,0,0,0,0,0,0,"
                Case 2 : FuncName = "SPU3ｷｬｯｼｭﾃｽﾄｲﾆｼｬﾗｲｻﾞ起動ｺﾏﾝﾄﾞ1" : SendCom = "TS,40,0,0,0,55555555," : RecvCom = "TS,40,0,0,0,0,"
                Case 3 : FuncName = "SPU3ｷｬｯｼｭﾃｽﾄｲﾆｼｬﾗｲｻﾞ起動ｺﾏﾝﾄﾞ2" : SendCom = "TS,40,0,0,0,AAAAAAAA," : RecvCom = "TS,40,0,0,0,0,"
                Case 4 : FuncName = "SPU3内部RAMﾃｽﾄ" : SendCom = "TS,41,0,0,C8000000,C8001FFF," : RecvCom = "TS,41,0,0,0,C8001FFC,0,0,"
            End Select
            Call WriteLog(False, FuncName & "中・・・")

            Try
                rcv = TSSendRecv(SendCom, 30)
                If i = 0 Then
                    If InStr(rcv, "Time Out") <> 0 Then
                        mErrorComment = FuncName & "　異常 " & rcv
                        Throw New ApplicationException(mErrorComment)
                    End If
                    MeasureValue = CInt("&H" & Strings.Mid(rcv, Len(RecvCom) + 1, 4))
                    If MeasureValue > CInt("&HAAA0") And MeasureValue < CInt("&HAE60") Then
                        Call WriteLog(False, "範囲内でOK。AAA0～AE60")
                    Else
                        mErrorComment = "SPUｸﾛｯｸﾃｽﾄ範囲外 AAA0～AE60：" & Strings.Mid(rcv, Len(RecvCom) + 1, 4)
                        mTestSpec = "基準：AAA0～AE60" 'Rev0.20add
                        Throw New ApplicationException(mErrorComment)
                    End If
                Else
                    If rcv <> RecvCom Then
                        mErrorComment = FuncName & "　異常 " & rcv
                        mTestSpec = RecvCom   'Rev0.20add
                        Throw New ApplicationException(mErrorComment)
                    End If
                End If
            Catch ex As Exception
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try
            Call WriteLog(False, FuncName & " OK ")

        Next i

    End Sub

    Public Sub I10(ByVal iName As String)
        '==============================================================================================================
        ' Function : I10:ﾗﾀﾞｰﾌﾗｯｼｭ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String
        Dim ReadWaitTime As Integer = 0
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""
        Dim Model As String = ModelRead()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        FuncName = "ﾗﾀﾞｰﾌﾗｯｼｭﾃｽﾄ" : SendCom = "TS,24,0,0,7,0,0,"
        If Strings.Left(Model, 6) = "F3SP76" Then RecvCom = "TS,24,0,0,0,7,30FFFFFC,0,0," : ReadWaitTime = 180 : Call WriteLog(False, FuncName & "中・・・約1分30秒")
        If Strings.Left(Model, 6) = "F3SP71" Then RecvCom = "TS,24,0,0,0,7,307FFFFC,0,0," : ReadWaitTime = 60 : Call WriteLog(False, FuncName & "中・・・約40秒")
        If Strings.Left(Model, 6) = "F3SPV9" Then RecvCom = "TS,24,0,0,0,7,30FFFFFC,0,0," : ReadWaitTime = 180 : Call WriteLog(False, FuncName & "中・・・約1分30秒") 'Rev0.12ADD
        Try
            Call CountStart(200)
            'rcv = TSSendRecv_LongWait(SendCom, ReadWaitTime)
            rcv = TSSendRecv(SendCom, ReadWaitTime)
            If rcv <> RecvCom Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = RecvCom    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
        End Try

        Call WriteLog(False, FuncName & " OK ")

    End Sub

    Public Sub I11(ByVal iName As String)
        '==============================================================================================================
        ' Function : I11:割込み信号ﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        FuncName = "割込み信号ﾃｽﾄ" : SendCom = "TS,31,0,0,8," : RecvCom = "TS,31,0,0,0,0,"
        Call WriteLog(False, FuncName & "中・・・")

        Try
            rcv = TSSendRecv(SendCom, 30)
            If rcv <> RecvCom Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = RecvCom    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, FuncName & " OK ")

    End Sub

    Public Sub I12(ByVal iName As String)
        '==============================================================================================================
        ' Function : I12:ｸﾛｯｸﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""
        Dim MeasureValue As Integer

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        'ﾃｽﾄｺﾏﾝﾄﾞ未定！
        FuncName = "ｸﾛｯｸﾃｽﾄ" : SendCom = "TS,20,0,0,0,0,78E0,7B40," : RecvCom = "TS,20,0,0,0,0,"
        Call WriteLog(False, FuncName & "中・・・")

        Try
            rcv = TSSendRecv(SendCom, 30)
            MeasureValue = CInt("&H" & Strings.Mid(rcv, RecvCom.Length + 1, 4))
            If MeasureValue > CInt("&H78E0") And MeasureValue < CInt("&H7B40") Then
                Call WriteLog(False, "範囲内でOK。78E0～7B40")
            Else
                mErrorComment = "範囲外。78E0～7B40　" & Strings.Mid(rcv, RecvCom.Length + 1, 4)
                mTestSpec = "基準値：78E0～7B40"    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WaitTimer(1)

        Call WriteLog(False, FuncName & " OK ")

    End Sub

    Public Sub I13(ByVal iName As String)
        '==============================================================================================================
        ' Function : I13:RTCﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim rcv As String
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        FuncName = "RTCﾃｽﾄ" : SendCom = "TS,8,0,0," : RecvCom = "TS,8,0,0,0,0,"
        Call WriteLog(False, FuncName & "中・・・")

        Try
            rcv = TSSendRecv(SendCom, 60)
            If rcv <> RecvCom Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = RecvCom    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WaitTimer(1)

        Call WriteLog(False, FuncName & " OK ")

    End Sub

    Public Sub I14(ByVal iName As String)
        '==============================================================================================================
        ' Function : I14:電源電圧検査
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
        Call WriteLog(False, "電源電圧検査中・・・")

        Dim ADDevice As String = ""
        Dim RecvDataLE As String = ""
        'Dim Result As String
        'Dim sValue As String
        Dim dValue As Single
        Dim StdMax As Single
        Dim StdMin As Single
        Dim SigName As String = ""

        Dim cMeas As New cls140_VoltageMeasure

        Dim TestStd_VCC_H As Single = 5.15
        Dim TestStd_VCC_L As Single = 4.85
        Dim TestStd_VCCSDI_H As Single = 3.4
        Dim TestStd_VCCSDI_L As Single = 3.0
        '------------------------ Rev 0.47 ATY DEL START ----------------------
        'Dim TestStd_VCC3_H As Single = 3.37
        'Dim TestStd_VCC3_L As Single = 3.23
        'Dim TestStd_V1P2_H As Single = 1.23
        'Dim TestStd_V1P2_L As Single = 1.17
        'Dim TestStd_V1P5_H As Single = 1.53
        'Dim TestStd_V1P5_L As Single = 1.47
        '------------------------ Rev 0.47 ATY DEL END ----------------------
        '------------------------ Rev 0.47 ATY ADD START ----------------------
        Dim TestStd_VCC3_H As Single = 3.4
        Dim TestStd_VCC3_L As Single = 3.2
        Dim TestStd_V1P2_H As Single = 1.24
        Dim TestStd_V1P2_L As Single = 1.16
        Dim TestStd_V1P5_H As Single = 1.55
        Dim TestStd_V1P5_L As Single = 1.46
        '------------------------ Rev 0.47 ATY ADD END ----------------------

        'cLnk = New cls080_F3LinkToLE
        'cLnk.IPAdd = IP_U1
        Try
            For i = 1 To 4
                Select Case i
                    Case 1
                        SigName = "VCC" : ADDevice = "0006" : StdMax = TestStd_VCC_H : StdMin = TestStd_VCC_L
                        'Case 2
                        '    SigName = "VCCSDI" : ADDevice = "0001" : StdMax = TestStd_VCCSDI_H : StdMin = TestStd_VCCSDI_L
                    Case 2
                        SigName = "VCC3P3" : ADDevice = "0002" : StdMax = TestStd_VCC3_H : StdMin = TestStd_VCC3_L
                    Case 3
                        SigName = "V1P2" : ADDevice = "0004" : StdMax = TestStd_V1P2_H : StdMin = TestStd_V1P2_L
                    Case 4
                        SigName = "V1P5" : ADDevice = "0005" : StdMax = TestStd_V1P5_H : StdMin = TestStd_V1P5_L
                End Select

                Call WriteLog(False, SigName & "検査中・・・" & "ch" & Strings.Right(ADDevice, 1))

                'Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
                'RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
                'Result = Strings.Left(RecvDataLE, 2)
                'If Result <> "OK" Then
                '    mErrorComment = SigName & "測定返信 異常 " & RecvDataLE
                '    Throw New ApplicationException(mErrorComment)
                'End If
                'sValue = Mid(RecvDataLE, 3, 4)
                'dValue = CDec("&h" & sValue) / (20000 / 10)

                'v0.07　DMM測定
                dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

                Call WriteLog(False, "AD測定値 : " & dValue.ToString & "V")
                If dValue > StdMax Or dValue < StdMin Then
                    mErrorComment = SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
                    mTestMeasure = dValue.ToString     'Rev0.20add
                    mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)

                Call WriteValueData(SigName, dValue.ToString)

            Next
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            'cLnk = Nothing
            cMeas = Nothing
        End Try

    End Sub

    Public Sub I15(ByVal iName As String)
        '==============================================================================================================
        ' Function : I15:Ethernet通信ﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        'Dim rcv As String
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
        Dim adapter As NetworkInterface

        Dim sKEY As String
        Dim NetAdpMAC As String
        Dim Ret As Boolean
        Dim Ret1 As Boolean
        Dim RegStr As String
        Dim RecvDataLE As String = ""
        Dim Test10M As Boolean = False
        Dim i As Integer = 0

        'ﾈｯﾄﾜｰｸｱﾀﾞﾌﾟﾀ確認 --------------------------------------------------------------------
        For Each adapter In adapters
            NetAdpMAC = adapter.GetPhysicalAddress.ToString
        Next adapter

        '設定 --------------------------------------------------------------------
        'LAN Speed Control
        sKEY = "*SpeedDuplex"       '←Win7用。 WinXPはこっち→"ConnectionType"

        '速度10M Fullに変更 ------------------------------------------------------
        Call WriteLog(False, "通信速度10M Fullに変更中・・・")

        DoEvents()
retry1:

        RegStr = "2" '10M Full
        cPCSP = New cls060_PCNetSPEED

        Ret1 = cPCSP.VBSetRegistryString(RegStr, sKEY)
        Ret = cPCSP.DisEnableNetAdapter

        cPCSP = Nothing

        Call WriteLog(False, "速度変更　反映待機中・・・")
        'Call WaitStart(5)    ' Ver 0.29 JDI DEL

        'FTP
        Call WriteLog(False, "通信速度10M Fullでの通信確認中・・・")
        '-------------------------- Ver 0.29 JDI DEL Start ------------------------
        'Try
        '    Call I05(iName)     'ﾃｽﾄﾓｰﾄﾞ移行
        '    Call I03(iName)     'ﾓｼﾞｭｰﾙ名称取得
        '    Call I06(iName)     'ﾃｽﾄﾓｰﾄﾞ終了
        '    Test10M = True
        '    Call WriteLog(False, "通信速度10M Fullでの通信　OK")
        '    Call WriteLog(False, "")

        'Catch ex As Exception

        'End Try
        '--------------------------- Ver 0.29 JDI DEL End -------------------------

        '-------------------------- Ver 0.29 JDI ADD Start ------------------------
        For i = 0 To 10
            If My.Computer.Network.Ping("192.168.0.2") = False Then
                Call WaitStart(1)
            Else
                Call WriteLog(False, "10M FULL通信OK")
                Exit For
            End If
            If i = 10 Then
                mErrorComment = "10M PING 通信異常"
                mTestSpec = "Ping 10回以下"     'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
        Next i
        '--------------------------- Ver 0.29 JDI ADD End -------------------------
        'Call mSecWait(1000)

        '速度100M Fullに変更 -----------------------------------------------------
        Call WriteLog(False, "通信速度100M Fullに変更中・・・")

        RegStr = "4" '100M Full

        cPCSP = New cls060_PCNetSPEED

        Ret1 = cPCSP.VBSetRegistryString(RegStr, sKEY)
        Ret = cPCSP.DisEnableNetAdapter

        Call WriteLog(False, "速度変更　反映待機中・・・")

        'Call WaitStart(5)             Rev0.12DEL

        '-----< Rev0.12ADD >-----
        For i = 0 To 10
            If My.Computer.Network.Ping("192.168.0.2") = False Then
                Call WaitStart(1)
            Else
                Call WriteLog(False, "100M FULL通信OK")
                Exit For
            End If
            If i = 10 Then
                mErrorComment = "100M PING 通信異常"
                mTestSpec = "Ping 10回以下"     'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
        Next i
        '------------------------

        ''FTP
        'Call WriteLog(False, "通信速度100M Fullでの通信確認中・・・")
        ''何で確認をするのか未定！

        'Call WriteLog(False, "通信速度100M Fullでの通信　OK")
        '-------------------------- Ver 0.29 JDI DEL Start ------------------------
        'If Test10M = False Then
        '    mErrorComment = "10M FULL通信異常"
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End If
        '--------------------------- Ver 0.29 JDI DEL End -------------------------

        'v0.07追加SDｶｰﾄﾞﾗﾀﾞｰ書き込み
        Call WriteLog(False, "SDｶｰﾄﾞﾗﾀﾞｰﾌﾟﾛｸﾞﾗﾑﾌｧｲﾙFTPｺﾋﾟｰ")
        Dim cSP6x As cls2D1_F3SP6x_Inspection
        cSP6x = New cls2D1_F3SP6x_Inspection
        Try
            If cSP6x.ReWriteSD() = False Then
                mErrorComment = "Ethernet通信検査(FTP)SD書換え 異常"
                mTestSpec = "SD書換え"   'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "Ethernet通信検査(FTP)SD書換え 異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP6x = Nothing
        End Try
        Call WriteLog(False, "SDｶｰﾄﾞﾗﾀﾞｰﾌﾟﾛｸﾞﾗﾑﾌｧｲﾙFTPｺﾋﾟｰOK")

        '正常終了 ----------------------------------------------------------------
        Call WriteLog(False, "Ethernet通信検査 完了")

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I16(ByVal iName As String)
        '==============================================================================================================
        ' Function : I16:ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟﾃｽﾄ 
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟ検査中・・・")

        Dim Model As String = ModelRead()
        Dim ADDevice As String
        Dim RecvDataLE As String = ""
        Dim Result As String
        'Dim sValue As String
        Dim dValue As Single
        Dim StdMax As Single
        Dim StdMin As Single
        Dim SigName As String
        Dim rcv As String
        Dim FuncName As String
        Dim PwrONWait As Integer
        Dim PwrOFFWait As Integer
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim i As Integer
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Dim TestStd_VC3SLP_OFF_H As Single = 3.15
        Dim TestStd_VC3SLP_OFF_L As Single = 2.5

        'Dim ResultData_DCVC3SLP_ON As Single
        'Dim ResultData_DCVC3SLP_OFF As Single
        'Dim ResultData_DCD_VC3SLP As Single
        'Dim ResultData_M1D_VC3SLP As Double
        Dim ResultData_VC3SLP_ON As Single
        'Dim ResultData_M1VC3SLP_OFF As Double
        'Dim ResultData_M0VC3SLP_OFF As Single
        'Dim ResultData_M0D_VC3SLP As Single
        Dim ResultData_VC3SLP_ON_FFFF As Single
        Dim ResultData_VC3SLP_OFF_FFFF As Single
        'Dim ResultData_VC3SLP_ON_0000 As Single                     ' Ver 0.28 JDI DEL
        'Dim ResultData_VC3SLP_OFF_0000 As Single                    ' Ver 0.28 JDI DEL

        Dim TestStd_VC3SLP_ON_H As Single
        Dim TestStd_VC3SLP_ON_L As Single

        Dim mV3P3 As Single

        Dim cMeas As New cls140_VoltageMeasure

        PwrONWait = 5
        PwrOFFWait = 3

        'Call I90("1")       'ｽﾛｯﾄ1で電源ON     Ver 0.28 JDI DEL


        '--------------------------- Ver 0.28 JDI DEL Start ---------------------------
        'Try
        '    'ﾘﾚｰ設定（GND）
        '    cLnk = New cls080_F3LinkToLE
        '    cLnk.IPAdd = IP_U1

        '    'VCC3P3測定
        '    ADDevice = "0002"
        '    'v0.07　DMM測定
        '    mV3P3 = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        '    TestStd_VC3SLP_ON_H = CSng(mV3P3 + 0.005)
        '    TestStd_VC3SLP_ON_L = CSng(mV3P3 - 0.05)

        '    'VC3SLP -------------------------------------
        '    SigName = "VC3SLP"
        '    ADDevice = "0003"
        '    StdMax = TestStd_VC3SLP_ON_H : StdMin = TestStd_VC3SLP_ON_L
        '    'v0.07　DMM測定
        '    dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        '    Call WriteLog(False, "電源ON時" & SigName & "測定中・・・")

        '    If dValue > StdMax Or dValue < StdMin Then
        '        mErrorComment = "電源ON時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
        '        mTestMeasure = dValue.ToString     'Rev0.20add
        '        mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
        '        Call WriteLog(False, mErrorComment)
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    Call WriteLog(False, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString)

        'Catch ex As Exception
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "501" & ",016,0000000000000000", RecvDataLE)
        '    cLnk = Nothing
        'End Try

        '---------------------------- Ver 0.28 JDI DEL End ----------------------------
        ResultData_VC3SLP_ON = dValue

        'RTC設定 -------------------------------------
        Call WriteLog(True, "RTC設定")      ' Ver 0.28 JDI ADD

        'Call WriteLog(False, "設定開始待ち・・・3秒")      ' Ver 0.28 KJ DEL
        'Call WaitStart(3)                                  ' Ver 0.28 KJ DEL

        'Ver 0.28 改善後NG発生によりWait追加 (10月/12日)
        'Ver 0.28 2秒でのNG発生は完全に解消できなかったので3秒に変更 (10月/23日)
        'Ver 0.28 RTC設定NG発生。待ち時間3秒->5秒に変更(Debug1月23日16:05) 

        'Call WriteLog(False, "設定開始待ち・・・5秒")      ' Ver 0.28 KJ ADD  ' Ver 0.35 JDI DEL
        'Call WaitStart(5)                                  ' Ver 0.28 KJ ADD  ' Ver 0.35 JDI DEL
        Call WriteLog(False, "設定開始待ち・・・3秒")      ' Ver 0.35 JDI ADD
        Call WaitStart(3)                                  ' Ver 0.35 JDI ADD

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_SP : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 0 : cMcm.TimeOut = 5
        i = 0
        Do
            Call cMcm.RTC_SET()     'RTC書き込み(現時刻)
            RecvDataLE = cMcm.RecvText
            If RecvDataLE = "0100021200020000" Then Exit Do

            If i >= 1 Then
                mErrorComment = iName & "RTC設定　返信　異常 " & RecvDataLE
                Call WriteLog(False, mErrorComment)
                mTestSpec = "0100021200020000"    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
            i = i + 1
            Call WriteLog(False, "RTC設定 ﾘﾄﾗｲ待ち　5秒")
            Call WaitStart(5)
            DoEvents()
        Loop

        cMcm = Nothing

        'Call WriteLog(False, "RTC設定 OK" & RecvDataLE)      ' Ver 0.28 JDI DEL
        Call WriteLog(True, "RTC設定 OK" & RecvDataLE)        ' Ver 0.28 JDI ADD
        Call mSecWait(500)

        '--------------------
        FuncName = "ﾃｽﾄﾓｰﾄﾞ移行"
        'Call WriteLog(False, FuncName & "・・・")            ' Ver 0.28 JDI DEL
        Call WriteLog(True, FuncName & "・・・")              ' Ver 0.28 JDI ADD

        Call I05(iName)

        'Call WriteLog(False, FuncName & " " & PwrOFFWait & "秒待ち")     ' Ver 0.28 JDI DEL
        'Call WaitStart(PwrOFFWait)                                       ' Ver 0.28 JDI DEL
        'Call WriteLog(False, "ﾃｽﾄﾓｰﾄﾞ移行　完了")            ' Ver 0.28 JDI DEL
        Call WriteLog(True, "ﾃｽﾄﾓｰﾄﾞ移行　完了")              ' Ver 0.28 JDI ADD

        '--------------------------- Ver 0.28 JDI ADD Start ---------------------------
        Dim testcommand As String

        Call WriteLog(False, "")
        Call WriteLog(False, "Test Mode Check")
        testcommand = TSSendRecv("TS,6,0,0,39046EA0,4,", 5)
        Call WriteLog(False, "All Clear Mode(2)→Testmode(3)に変更")
        testcommand = TSSendRecv("TS,7,0,0,39046EA0,4,3,", 5)
        Call WriteLog(False, "Test Mode Check")
        testcommand = TSSendRecv("TS,6,0,0,39046EA0,4,", 5)
        Call WriteLog(False, "")
        '---------------------------- Ver 0.28 JDI ADD END ----------------------------

        ''準備：KEYｵｰﾌﾟﾝ
        'SendCom = "TS,7,0,0,39800F00,4,53505533," : RecvCom = "TS,7,0,0,1,4,39800F00,0,"
        'rcv = TSSendRecv(SendCom, 5)

        'v0.07検査手順変更
        'ﾊﾞｯｸｱｯﾌﾟSRAM全領域に$FFFFを書き込み
        'If Strings.Left(Model, 6) = "F3SP76" Then SendCom = "TS,21,0,0,3B400000,3B5FFFFF,2,FFFF," 'Ver 0.28 JDI DEL
        'If Strings.Left(Model, 6) = "F3SP71" Then SendCom = "TS,21,0,0,3B400000,3B4FFFFF,2,FFFF," 'Ver 0.28 JDI DEL
        'If Strings.Left(Model, 6) = "F3SPV9" Then SendCom = "TS,21,0,0,3B400000,3B5FFFFF,2,FFFF," 'Rev0.12ADD 'Ver 0.28 JDI DEL

        '検査時間短縮のため、SRAM全領域に$FFFFを書き込みから、SRAMのバックアップ領域だけを確認する方法に変更
        If Strings.Left(Model, 6) = "F3SP76" Then SendCom = "TS,21,0,0,3907B400,390FB3FC,4,AA55AA55," 'Ver 0.28 JDI ADD
        If Strings.Left(Model, 6) = "F3SP71" Then SendCom = "TS,21,0,0,3907B400,3908B3FC,4,AA55AA55," 'Ver 0.28 JDI ADD
        If Strings.Left(Model, 6) = "F3SPV9" Then SendCom = "TS,21,0,0,3907B400,390FB3FC,4,AA55AA55," 'Ver 0.28 JDI ADD


        rcv = TSSendRecv(SendCom, 5)
        'Call WaitTimer(1)       'Ver 0.28 JDI DEL

        Call WriteLog(True, "SRAMのﾊﾞｯｸｱｯﾌﾟ領域に書き込み完了。")       'Ver 0.28 JDI ADD

        Try
            'ﾘﾚｰ設定（GND）
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U1
            '--------------------------- Ver 0.29 JDI DEL Start ---------------------------
            '--------------------------- Ver 0.28 JDI ADD Start ---------------------------
            ''VCC3P3測定
            'ADDevice = "0002"
            ''v0.07　DMM測定
            'mV3P3 = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

            'TestStd_VC3SLP_ON_H = CSng(mV3P3 + 0.005)
            'TestStd_VC3SLP_ON_L = CSng(mV3P3 - 0.05)
            '---------------------------- Ver 0.28 JDI ADD END ----------------------------
            '---------------------------- Ver 0.29 JDI DEL END ----------------------------

            '--------------------------- Ver 0.29 JDI ADD Start ---------------------------
            Dim dbName As String = ""
            Dim tbName As String = ""
            Dim sSQL As String = ""
            Dim rsData() As String = {""}
            Dim rdData() As String = {""}
            Dim rfName() As String = {""}
            Dim rdName() As String = {""}

            Dim DataCount As Integer = 0

            dbName = mAppPath & "210 Result Data\Temporary.mdb"
            tbName = "tblResult"
            sSQL = "SELECT StartDateTime,MSCODE,SerialNo FROM " & tbName

            Try
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(rfName, rsData)

                DataCount = cADO.DB_DATA_COUNT

            Catch ex As Exception
                mErrorComment = cADO.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString
                Throw New ApplicationException(mErrorComment)
            Finally
                cADO = Nothing
            End Try

            If DataCount = 0 Then
                ADDevice = "0002"

                mV3P3 = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))
                Call WriteLog(True, "DMM V3P3 : " & mV3P3 & "V")

                TestStd_VC3SLP_ON_H = CSng(mV3P3 + 0.005)
                TestStd_VC3SLP_ON_L = CSng(mV3P3 - 0.05)
            Else
                dbName = mAppPath & "210 Result Data\F3SP7XData.mdb"
                tbName = "tblResult"

                'データがあるかどうかチェック
                sSQL = "SELECT * FROM " & tbName & " "
                sSQL &= "WHERE 検査日時='" & rsData(0) & "' AND MSCODE='" & rsData(1) & "' AND SerialNo='" & rsData(2) & "'"
                Try
                    cADO = New cls010_MDB_Operation
                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName
                    cADO.DB_SQL = sSQL
                    Call cADO.DB_READ_JET(rdName, rdData)
                    DataCount = cADO.DB_DATA_COUNT
                Catch ex As Exception
                    mErrorComment = cADO.ERR_COMMENT
                    If mErrorComment = "" Then mErrorComment = ex.ToString
                    Throw New ApplicationException(mErrorComment)
                Finally
                    cADO = Nothing
                End Try

                If DataCount = 0 Then
                    ADDevice = "0002"

                    mV3P3 = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))
                    Call WriteLog(True, "DMM V3P3 : " & mV3P3 & "V")

                    TestStd_VC3SLP_ON_H = CSng(mV3P3 + 0.005)
                    TestStd_VC3SLP_ON_L = CSng(mV3P3 - 0.05)
                Else
                    Call WriteLog(True, "DB V3P3 : " & rdData(4) & "V")
                    TestStd_VC3SLP_ON_H = CSng(CDbl(rdData(4)) + 0.005)
                    TestStd_VC3SLP_ON_L = CSng(CDbl(rdData(4)) - 0.05)
                End If

            End If
            '---------------------------- Ver 0.29 JDI ADD END ----------------------------

            'VC3SLP -------------------------------------
            SigName = "VC3SLP"
            ADDevice = "0003"
            StdMax = TestStd_VC3SLP_ON_H : StdMin = TestStd_VC3SLP_ON_L
            'v0.07　DMM測定
            dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

            Call WriteLog(False, "電源ON時" & SigName & "測定中・・・")

            If dValue > StdMax Or dValue < StdMin Then
                mErrorComment = "電源ON時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
                mTestMeasure = dValue.ToString     'Rev0.20add
                mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If

            'Call WriteLog(False, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString)  Ver 0.28 JDI DEL
            Call WriteLog(True, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString)  'Ver 0.28 JDI ADD

        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        Finally
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "501" & ",016,0000000000000000", RecvDataLE)
            cLnk = Nothing
        End Try

        ResultData_VC3SLP_ON_FFFF = dValue

        '電源OFF
        Call I91(iName)
        'Call WriteLog(False, "電源遮断 " & PwrOFFWait & "秒待ち")  ' Ver 0.28 JDI DEL
        'Call WaitStart(PwrOFFWait)                                 ' Ver 0.28 JDI DEL

        '--------------------- Ver 0.28 JDI ADD Start --------------------------
        '放電抵抗の接続 -------------------------------------
        Call WriteLog(False, "")
        Call WriteLog(True, "放電抵抗(330kohm)接続")

        Try

            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U1

            'YC16-0Nのch11
            Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "511" & ",001,1", RecvDataLE)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(True, "1秒間放電抵抗接続")
            Call WaitStart(1)

            
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        Finally

            'ﾘﾚｰ設定OFF
            Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "511" & ",001,0", RecvDataLE)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If

            cLnk = Nothing
        End Try

        Call WaitStart(2)
        '---------------------- Ver 0.28 JDI ADD END ---------------------------

        'VC3SLP -------------------------------------
        SigName = "VC3SLP"
        ADDevice = "0003"
        StdMax = TestStd_VC3SLP_OFF_H : StdMin = TestStd_VC3SLP_OFF_L

        Call WriteLog(False, "電源OFF時" & SigName & "測定中・・・")

        Try
            'v0.07　DMM測定
            dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            'cLnk = Nothing
        End Try

        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = "電源OFF時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
            mTestMeasure = dValue.ToString     'Rev0.20add
            mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "電源OFF時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString("0.00"))
        ResultData_VC3SLP_OFF_FFFF = dValue

        '電源投入 -------------------------------------
        Call WriteLog(True, "電源投入開始")   'Ver 0.28 JDI ADD
        Call I90("1")

        Call WriteLog(True, "電源投入完了")   'Ver 0.28 JDI ADD

        'Call WriteLog(True, "電源投入 " & PwrONWait & "秒待ち")   'Ver 0.28 JDI DEL

        'Call WaitStart(PwrONWait)                                 'Ver 0.28 JDI DEL

        '--------------------- Ver 0.28 JDI ADD Start --------------------------
        Call I05(iName)

        Call WriteLog(True, "テストモード移行完了。")

        'ﾊﾞｯﾃﾘ異常ﾘﾚｰ確認
        Call WriteLog(False, "")
        Call WriteLog(False, Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "＜判定①＞ﾊﾞｯﾃﾘ異常ﾘﾚｰM194確認")

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U1
            Call cLnk.F3PCLinkCommunication("01" & "BRD" & "M194" & ",001", RecvDataLE)

            Result = Strings.Mid(RecvDataLE, 3, 2)
            If Result <> "OK" Then
                mErrorComment = "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194　異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If
            If Mid(RecvDataLE, 5, 1) <> "0" Then
                mErrorComment = "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194異常 " & Mid(RecvDataLE, 3, 4) & " (正常は0)"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194=0確認 OK")

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        'SRAMのMemory確認
        Dim rcv1 As String
        Dim rcv2 As String
        Dim StartAddress As Integer
        Dim FileReg As Integer
        Dim RegCount As Integer
        Dim NGCount As Integer
        Dim FileRegStep As Integer

        Call WriteLog(False, "")
        Call WriteLog(True, "＜判定②＞SRAMのMemory確認")

        StartAddress = 956806144 '$3907 B400

        If Strings.Left(Model, 6) = "F3SP76" Or Strings.Left(Model, 6) = "F3SPV9" Then
            FileReg = 262143
            FileRegStep = 300
        ElseIf Strings.Left(Model, 6) = "F3SP71" Then
            FileReg = 32767
            FileRegStep = 38
        End If

        For i = 0 To FileReg Step FileRegStep
            RegCount = RegCount + 1
            rcv2 = (StartAddress + i).ToString("X")
            rcv1 = TSSendRecv_NonDisp("TS,6,0,0," & rcv2 & ",4,", 5)

            If Not Strings.Mid(rcv1, 14, 8) = "AA55AA55" Then
                NGCount = NGCount + 1
            End If
        Next

        Call WriteLog(True, "レジスタ Count : " & RegCount)
        Call WriteLog(False, "NG Count : " & NGCount)
        Call WriteLog(False, "NG 発生率 : " & (NGCount / RegCount) * 100 & "%")

        If NGCount > 0 Then
            mErrorComment = "SRAMのMemory確認 "
            Call WriteLog(True, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        '---------------------- Ver 0.28 JDI ADD END ---------------------------

        '--------------------- Ver 0.28 JDI DEL Start --------------------------
        'FuncName = "ﾃｽﾄﾓｰﾄﾞ移行"
        'Call WriteLog(False, FuncName & "・・・")
        'Call I05(iName)

        ''ﾊﾞｯｸｱｯﾌﾟSRAM全領域に0を書き込み
        'If Strings.Left(Model, 6) = "F3SP76" Then SendCom = "TS,21,0,0,3B400000,3B5FFFFF,2,0000,"
        'If Strings.Left(Model, 6) = "F3SP71" Then SendCom = "TS,21,0,0,3B400000,3B4FFFFF,2,0000,"
        'If Strings.Left(Model, 6) = "F3SPV9" Then SendCom = "TS,21,0,0,3B400000,3B4FFFFF,2,0000," 'Rev0.12ADD
        'rcv = TSSendRecv(SendCom, 5)
        'Call WaitTimer(1)

        ''ﾃｽﾄﾓｰﾄﾞ終了
        'Call I06(iName)

        ''電源OFF
        'Call I91(iName)

        'Call WriteLog(False, "電源遮断 " & PwrOFFWait & "秒待ち")
        'Call WaitStart(PwrOFFWait)

        ''VC3SLP -------------------------------------
        'SigName = "VC3SLP"
        'ADDevice = "0003"
        'StdMax = TestStd_VC3SLP_OFF_H : StdMin = TestStd_VC3SLP_OFF_L

        'Call WriteLog(False, "電源OFF時" & SigName & "測定中・・・")

        'cLnk = New cls080_F3LinkToLE
        'Try
        '    'v0.07　DMM測定
        '    dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    cLnk = Nothing
        'End Try

        'If dValue > StdMax Or dValue < StdMin Then
        '    mErrorComment = "電源OFF時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
        '    mTestMeasure = dValue.ToString     'Rev0.20add
        '    mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End If

        'Call WriteLog(False, "電源OFF時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString("0.00"))
        'ResultData_VC3SLP_OFF_0000 = dValue

        ' ''VC3SLP差分 -------------------------------------
        ''ResultData_M0D_VC3SLP = ResultData_VC3SLP_ON - ResultData_M0VC3SLP_OFF
        ''Call WriteLog(False, "ON時電圧-OFF時電圧=" & CStr(ResultData_M0D_VC3SLP))
        ''Call WriteLog(False, "")


        ''OFF時　Wait -------------------------------------
        'Call WriteLog(False, "")
        ''Call WriteLog(False, "放電待ち 1分30秒待ち")
        'Call WriteLog(False, "15秒電源OFF")
        'Call WriteLog(False, "")

        'Call WaitStart(15)

        ''電源投入 -------------------------------------
        'Call WriteLog(False, "電源投入")

        'Call I90("1")

        'Try
        '    'VC3SLP -------------------------------------
        '    SigName = "VC3SLP"
        '    ADDevice = "0003"
        '    StdMax = TestStd_VC3SLP_ON_H : StdMin = TestStd_VC3SLP_ON_L

        '    'v0.07　DMM測定
        '    dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        '    Call WriteLog(False, "電源ON時" & SigName & "測定中・・・")

        '    If dValue > StdMax Or dValue < StdMin Then
        '        mErrorComment = "電源ON時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
        '        mTestMeasure = dValue.ToString     'Rev0.20add
        '        mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
        '        Call WriteLog(False, mErrorComment)
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    Call WriteLog(False, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString)

        '    'ResultData_VC3SLP_ON = dValue

        'Catch ex As Exception
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "501" & ",016,0000000000000000", RecvDataLE)
        '    cLnk = Nothing
        'End Try

        ''Call WriteLog(False, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString)

        'ResultData_VC3SLP_ON_0000 = dValue

        ' ''VC3SLP差分() -------------------------------------
        ''ResultData_DCD_VC3SLP = ResultData_DCVC3SLP_ON - ResultData_DCVC3SLP_OFF
        ''Call WriteLog(False, "ON時電圧-OFF時電圧=" & CStr(ResultData_DCD_VC3SLP))
        ''Call WriteLog(False, "")

        'Call WaitStart(3)

        ''--------------------
        ''ﾊﾞｯﾃﾘ異常ﾘﾚｰ確認
        'Call WriteLog(False, "")
        'Call WriteLog(False, "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194確認")

        'Try
        '    cLnk = New cls080_F3LinkToLE
        '    cLnk.IPAdd = IP_U1
        '    Call cLnk.F3PCLinkCommunication("01" & "BRD" & "M194" & ",001", RecvDataLE)

        '    Result = Strings.Mid(RecvDataLE, 3, 2)
        '    If Result <> "OK" Then
        '        mErrorComment = "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194　異常 " & RecvDataLE
        '        mTestSpec = "正常:OK"   'Rev0.20add
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    If Mid(RecvDataLE, 5, 1) <> "0" Then
        '        mErrorComment = "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194異常 " & Mid(RecvDataLE, 3, 4) & " (正常は0)"
        '        mTestSpec = "正常:OK"   'Rev0.20add
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    Call WriteLog(False, "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194=0確認 OK")

        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    cLnk = Nothing
        'End Try
        '---------------------- Ver 0.28 JDI DEL END ---------------------------
    End Sub

    Public Sub I17(ByVal iName As String)
        '==============================================================================================================
        ' Function : I17:USB通信ﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim Model As String = ModelRead()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Dim usb As New COMCTRLEXLib.ComMng
        Dim oInfo As New COMCTRLEXLib.F3OpenInfo2

        Call WaitStart(3)

        Try
            'usb.PortSetWakeUp(0, 1)
            usb.Comopn2(0, oInfo, 1)
            Call WriteLog(False, oInfo.xCpu.strSysID.ToString)
            Call WriteLog(False, oInfo.xCpu.strRev.ToString)
            If Strings.Left(Model, 9) <> oInfo.xCpu.strSysID.ToString Then
                mErrorComment = "形名読み間違い"
                mTestSpec = Strings.Left(Model, 9)    'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            usb.Comcls(0)
        Catch ex As Exception
            mErrorComment = "USB通信異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            usb = Nothing
            oInfo = Nothing
        End Try

        Call WriteLog(False, "USB通信ﾃｽﾄOK")

    End Sub

    Public Sub I18(ByVal iName As String)
        '==============================================================================================================
        ' Function : I18:IOﾊﾞｽ-ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01
        Call WriteLog(False, "ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄ中・・・約5秒")
        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.SlotAddressTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄOK")

    End Sub

    Public Sub I19(ByVal iName As String)
        '==============================================================================================================
        ' Function : I19:IOﾊﾞｽ-XACSﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01
        Call WriteLog(False, "XACS中・・・約3秒")
        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.XACSTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "XACSﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "XACS検査OK")

    End Sub

    Public Sub I20(ByVal iName As String)
        '==============================================================================================================
        ' Function : I20:IOﾊﾞｽ-XBRﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01

        'XBRﾃｽﾄ
        Call WriteLog(False, "XBR検査中・・・約3秒")
        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.XBRTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "XBRﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "XBR検査OK")

    End Sub

    Public Sub I21(ByVal iName As String)
        '==============================================================================================================
        ' Function : I21:IOﾊﾞｽ-XDFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim mResStr As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01

        Call WriteLog(False, "XDFAIL検査中・・・約10秒")

        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.XDFAILTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "XDFAILﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "XDFAIL検査OK")

    End Sub

    Public Sub I22(ByVal iName As String)
        '==============================================================================================================
        ' Function : I22:IOﾊﾞｽ-割込みﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim mResStr As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01

        Call WriteLog(False, "XIRQ検査中・・・約3秒")

        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.XIRQTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "XIRQﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "XIRQ検査OK")

    End Sub

    Public Sub I23(ByVal iName As String)
        '==============================================================================================================
        ' Function : I23:IOﾊﾞｽ-XLFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01

        Call WriteLog(False, "XLFAIL検査中・・・約1秒")

        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.XLFAILTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "XLFAILﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "XLFAIL検査OK")

    End Sub

    Public Sub I24(ByVal iName As String)
        '==============================================================================================================
        ' Function : I24:IOﾊﾞｽ-XPFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01

        Call WriteLog(False, "XPFAIL検査中・・・約1秒")

        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.XPFAILTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "XPFAILﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "XPFAIL検査OK")

    End Sub

    Public Sub I25(ByVal iName As String)
        '==============================================================================================================
        ' Function : I25:IOﾊﾞｽ-XSFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01

        Call WriteLog(False, "XSFAIL検査中・・・約1秒")

        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = 5
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cBUBU.XSFAILTest() = False Then
                Call WriteLog(False, cBUBU.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cBUBU.LOG)
        Catch ex As Exception
            mErrorComment = "XSFAILﾃｽﾄ異常:" & cBUBU.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "XSFAIL検査OK")

    End Sub

    Public Sub I26(ByVal iName As String)
        '==============================================================================================================
        ' Function : I26:LED検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""
        Dim rcv As String
        Dim fCam As Boolean = False
        Dim Model As String = ModelRead()
        'Dim res As MsgBoxResult

        Dim NG_Count As Integer           'Ver 0.24 JDI ADD
        Dim Test_Count As Integer         'Ver 0.24 JDI ADD

        Dim cEI As cls130_EI_GAZO

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        '------------------------------- Ver 0.27 JDI ADD Start -----------------------------------
        Dim mCommand_LED_ON(,) As String = {{"RDY소등", "TS,7,0,0,3880B21B,1,0,", "TS,7,0,0,0,1,3880B21B,0,"}, _
                                     {"RDY점등", "TS,7,0,0,3880B21B,1,C0,", "TS,7,0,0,1,1,3880B21B,0,"}, _
                                     {"RUN점등", "TS,7,0,0,3880B228,1,80,", "TS,7,0,0,0,1,3880B228,80,"}, _
                                     {"ALM점등", "TS,7,0,0,3880B229,1,80,", "TS,7,0,0,0,1,3880B229,80,"}, _
                                     {"ERR점등", "TS,7,0,0,3880B22A,1,80,", "TS,7,0,0,0,1,3880B22A,80,"}, _
                                     {"SD,US1점등", "TS,7,0,0,FFFE3A02,2,0FAF,", "TS,7,0,0,0,2,FFFE3A02,FAF,"}, _
                                     {"EXE,US2점등", "TS,7,0,0,FFFE3A02,2,0F5F,", "TS,7,0,0,0,2,FFFE3A02,F5F,"}}

        Dim mCommand_LED_OFF(,) As String = {{"RDY소등", "TS,7,0,0,3880B21B,1,0,", "TS,7,0,0,0,1,3880B21B,0,"}, _
                                     {"RUN소등", "TS,7,0,0,3880B228,1,0,", "TS,7,0,0,0,1,3880B228,0,"}, _
                                     {"ALM소등", "TS,7,0,0,3880B229,1,0,", "TS,7,0,0,0,1,3880B229,0,"}, _
                                     {"ERR소등", "TS,7,0,0,3880B22A,1,0,", "TS,7,0,0,0,1,3880B22A,0,"}, _
                                     {"SD,EXE,US1,US2소등", "TS,7,0,0,FFFE3A02,2,0FFF,", "TS,7,0,0,0,2,FFFE3A02,FFF,"}}

        '-------------------------------- Ver 0.27 JDI ADD End ------------------------------------


        '画像処理入れるときは、、、
        'V0.01画像処理をｸﾗｽﾓｼﾞｭｰﾙ化
        Try
            cEI = New cls130_EI_GAZO
            cEI.IPAdd = "192.168.1.30" : cEI.MasterFile = Mid(Model, 1, 9) & "_NonLED.jpg"
            If cEI.MasterFileRead = False Then
                fCam = False
                Throw New ApplicationException(mErrorComment)
            Else
                fCam = True
            End If
        Catch ex As Exception
            cEI = Nothing
        End Try

        NG_Count = 0                  'Ver 0.24 JDI ADD
        Test_Count = 0                'Ver 0.24 JDI ADD

        '------------------------- Ver 0.27 JDI ADD Start -----------------------
        rcv = TSSendRecv_NonDisp(mCommand_LED_ON(1, 1), 3)

        If rcv <> mCommand_LED_ON(1, 2) Then
            mErrorComment = mCommand_LED_ON(1, 0) & "RCV Data 異常"
            Throw New ApplicationException(mErrorComment)
        End If
        '-------------------------- Ver 0.27 JDI ADD End ------------------------

        Select Case fCam
            Case True       'カメラありのとき
                Call WriteLog(False, "LEDﾌｨﾙﾀﾁｪｯｸ")

                Test_Count += 1       'Ver 0.24 JDI ADD

                '------------------------- Ver 0.27 JDI DEL Start -----------------------
                'rcv = TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,0,", 3)       '全消灯
                ''------------------------- Ver 0.25 JDI ADD Start -----------------------
                'RecvCom = "TS,7,0,0,0,1,3880B21B,0,"

                'Call WriteLog(False, "全消灯 RCV:" & rcv)

                'If rcv <> RecvCom Then
                '    mErrorComment = "RCV Data 異常"
                '    Throw New ApplicationException(mErrorComment)
                'End If
                ''-------------------------- Ver 0.25 JDI ADD End ------------------------
                '-------------------------- Ver 0.27 JDI DEL End ------------------------
                Call mSecWait(500)
                Call cEI.GAZO_TEST()
                Call WriteLog(False, cEI.ResultData)
                If cEI.RESULT <> "0" Then
                    NG_Count += 1                                           'Ver 0.24 JDI ADD           
                    Call WriteLog(False, "ﾌｨﾙﾀﾁｪｯｸ目視検査")                'Ver 0.24 JDI ADD

                    'メッセージ表示
                    dfrm = New frm500_MessagePopUp
                    'dfrm.DISP_TITLENAME = "LEDﾌｨﾙﾀ検査"                                                'Rev0.11DEL
                    'dfrm.DISP_MESSAGE = "LEDﾌｨﾙﾀｼﾙｸ印刷は合っていますか？MODEL=" & Mid(Model, 1, 9)               'Rev0.11DEL
                    'dfrm.DISP_TITLENAME = "LED필터 검사"                                                'Rev0.11DEL
                    'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까? MODEL=" & Mid(Model, 1, 9)        'Rev0.11ADD     Ver 0.23 JDI DEL
                    'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까?" & vbCrLf & "MODEL=" & Mid(Model, 3, 7)        'Ver 0.23 ADD

                    dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7x\" & Mid(Model, 3, 4) & "_Filter.ppt"        'Ver 0.27 JDI ADD
                    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                    Call dfrm.frmShow()

                    Do
                        'wait push button -------------------
                        If dfrm.DIALOG_RESULT = "NG" Then
                            dfrm.Close()
                            dfrm = Nothing
                            mErrorComment = FuncName & "検査 異常 "
                            mTestSpec = "필터검사NG"     'Rev0.20add
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

                cEI.MasterFile = "F3SP76-7N_LED.jpg"
                Call cEI.MasterFileRead()

                '------------------------------------ Ver 0.27 JDI DEL Start -------------------------------------------
                '                For i = 0 To 15
                '                    Select Case i
                '                        '-----< Rev0.11DEL >-----
                '                        'Case 0 : FuncName = "全消灯" : SendCom = "TS,7,0,0,3880B21B,1,0," : RecvCom = "TS,7,0,0,0,1,3880B21B,0,"
                '                        'Case 1 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,3880B21B,1,C0," : RecvCom = "TS,7,0,0,1,1,3880B21B,0,"
                '                        'Case 2 : FuncName = "RDY,RUN点灯" : SendCom = "TS,7,0,0,3880B228,1,80," : RecvCom = "TS,7,0,0,0,1,3880B228,80,"
                '                        'Case 3 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,3880B228,1,0," : RecvCom = "TS,7,0,0,0,1,3880B228,0,"
                '                        'Case 4 : FuncName = "RDY,ALM点灯" : SendCom = "TS,7,0,0,3880B229,1,80," : RecvCom = "TS,7,0,0,0,1,3880B228,80,"
                '                        'Case 5 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,3880B229,1,0," : RecvCom = "TS,7,0,0,0,1,3880B229,0,"
                '                        'Case 6 : FuncName = "RDY,ERR点灯" : SendCom = "TS,7,0,0,3880B22A,1,80," : RecvCom = "TS,7,0,0,0,1,3880B22A,80,"
                '                        'Case 7 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,3880B22A,1,0," : RecvCom = "TS,7,0,0,0,1,3880B22A,0,"
                '                        'Case 8 : FuncName = "RDY,SD点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FEF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FEF,"
                '                        'Case 9 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                        'Case 10 : FuncName = "RDY,EXE点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FDF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FDF,"
                '                        'Case 11 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                        'Case 12 : FuncName = "RDY,US1点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FBF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FBF,"
                '                        'Case 13 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                        'Case 14 : FuncName = "RDY,US2点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0F7F," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,F7F,"
                '                        'Case 15 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                        Case 0 : FuncName = "전소등" : SendCom = "TS,7,0,0,3880B21B,1,0," : RecvCom = "TS,7,0,0,0,1,3880B21B,0,"
                '                        Case 1 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,3880B21B,1,C0," : RecvCom = "TS,7,0,0,1,1,3880B21B,0,"
                '                        Case 2 : FuncName = "RDY,RUN점등" : SendCom = "TS,7,0,0,3880B228,1,80," : RecvCom = "TS,7,0,0,0,1,3880B228,80,"
                '                        Case 3 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,3880B228,1,0," : RecvCom = "TS,7,0,0,0,1,3880B228,0,"
                '                            'Case 4 : FuncName = "RDY,ALM점등" : SendCom = "TS,7,0,0,3880B229,1,80," : RecvCom = "TS,7,0,0,0,1,3880B228,80,"         Ver 0.25 JDI DEL
                '                        Case 4 : FuncName = "RDY,ALM점등" : SendCom = "TS,7,0,0,3880B229,1,80," : RecvCom = "TS,7,0,0,0,1,3880B229,80,"             'Ver 0.25 JDI ADD
                '                        Case 5 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,3880B229,1,0," : RecvCom = "TS,7,0,0,0,1,3880B229,0,"
                '                        Case 6 : FuncName = "RDY,ERR점등" : SendCom = "TS,7,0,0,3880B22A,1,80," : RecvCom = "TS,7,0,0,0,1,3880B22A,80,"
                '                        Case 7 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,3880B22A,1,0," : RecvCom = "TS,7,0,0,0,1,3880B22A,0,"
                '                        Case 8 : FuncName = "RDY,SD점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FEF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FEF,"
                '                        Case 9 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                        Case 10 : FuncName = "RDY,EXE점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FDF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FDF,"
                '                        Case 11 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                        Case 12 : FuncName = "RDY,US1점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FBF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FBF,"
                '                        Case 13 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                        Case 14 : FuncName = "RDY,US2점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0F7F," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,F7F,"
                '                        Case 15 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
                '                            '-----------------------
                '                    End Select

                '                    Test_Count += 1       'Ver 0.24 JDI ADD

                '                    Try
                '                        rcv = TSSendRecv_NonDisp(SendCom, 3)
                '                        Call mSecWait(500)

                '                        '------------------------- Ver 0.25 JDI ADD Start -----------------------
                '                        Call WriteLog(False, FuncName & "RCV:" & rcv)

                '                        If rcv <> RecvCom Then
                '                            mErrorComment = "RCV Data 異常"
                '                            Throw New ApplicationException(mErrorComment)
                '                        End If
                '                        '-------------------------- Ver 0.25 JDI ADD End ------------------------

                '                    Catch ex As Exception
                '                        Throw New ApplicationException(mErrorComment)
                '                    End Try

                '                    Call cEI.GAZO_TEST()

                '                    Try
                '                        Select Case FuncName
                '                            '-----< Rev0.11DEL >-----
                '                            'Case "RDY点灯" : If cEI.RESULT <> "0,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RUN点灯" : If cEI.RESULT <> "1,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RDY,ALM点灯" : If cEI.RESULT <> "0,1,0,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RDY,ERR点灯" : If cEI.RESULT <> "0,1,1,0,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RDY,SD点灯" : If cEI.RESULT <> "0,1,1,1,0,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RDY,EXE点灯" : If cEI.RESULT <> "0,1,1,1,1,0,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RDY,US1点灯" : If cEI.RESULT <> "0,1,1,1,1,1,0,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RDY,US2点灯" : If cEI.RESULT <> "0,1,1,1,1,1,1,0,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "全消灯" : If cEI.RESULT <> "1,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            'Case "RDY,RUN点灯" : If cEI.RESULT <> "0,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY점등" : If cEI.RESULT <> "0,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RUN점등" : If cEI.RESULT <> "1,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY,ALM점등" : If cEI.RESULT <> "0,1,0,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY,ERR점등" : If cEI.RESULT <> "0,1,1,0,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY,SD점등" : If cEI.RESULT <> "0,1,1,1,0,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY,EXE점등" : If cEI.RESULT <> "0,1,1,1,1,0,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY,US1점등" : If cEI.RESULT <> "0,1,1,1,1,1,0,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY,US2점등" : If cEI.RESULT <> "0,1,1,1,1,1,1,0,1,1,1,1" Then GoTo ErrCheck
                '                            Case "전소등" : If cEI.RESULT <> "1,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                            Case "RDY,RUN점등" : If cEI.RESULT <> "0,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                '                                '-----------------------------
                '                        End Select

                '                        GoTo NextGo

                'ErrCheck:
                '                        NG_Count += 1                                   'Ver 0.24 JDI ADD
                '                        Call WriteLog(False, "目視検査")                'Ver 0.24 JDI ADD

                '                        'メッセージ表示
                '                        dfrm = New frm500_MessagePopUp
                '                        'dfrm.DISP_TITLENAME = "LED 検査"                         'Rev0.11DEL
                '                        'dfrm.DISP_MESSAGE = FuncName & "していますか？"          'Rev0.11DEL
                '                        dfrm.DISP_TITLENAME = "LED 검사"                         'Rev0.11ADD
                '                        dfrm.DISP_MESSAGE = FuncName & "하고 있습니까?"          'Rev0.11ADD
                '                        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                '                        Call dfrm.frmShow()

                '                        Do
                '                            'wait push button -------------------
                '                            If dfrm.DIALOG_RESULT = "NG" Then
                '                                dfrm.Close()
                '                                dfrm = Nothing
                '                                mErrorComment = FuncName & "検査 異常 "
                '                                mTestSpec = FuncName   'Rev0.20add
                '                                Throw New ApplicationException(mErrorComment)
                '                            End If

                '                            If dfrm.DIALOG_RESULT = "OK" Then
                '                                Exit Do
                '                            End If

                '                            DoEvents()

                '                        Loop

                '                        dfrm.Close()
                '                        dfrm = Nothing

                '                    Catch ex As Exception
                '                        cEI = Nothing
                '                        Call WriteLog(False, mErrorComment)
                '                        Throw New ApplicationException(mErrorComment)
                '                    End Try

                'NextGo:
                '                    Call WriteLog(False, FuncName & " OK")
                '                    Call WriteLog(False, cEI.RESULT)

                '                Next

                'cEI = Nothing
                '------------------------------------- Ver 0.27 JDI DEL End --------------------------------------------

                '------------------------------------ Ver 0.27 JDI ADD Start -------------------------------------------
                For i = 1 To 3
                    Select Case i
                        Case 1
                            Call WriteLog(False, "全消灯")
                            For j = 0 To 4

                                rcv = TSSendRecv_NonDisp(mCommand_LED_OFF(j, 1), 3)

                                If rcv <> mCommand_LED_OFF(j, 2) Then
                                    Call WriteLog(False, "RCV" & i & ":" & rcv)
                                    mErrorComment = mCommand_LED_ON(j, 0) & "RCV Data 異常"
                                    Throw New ApplicationException(mErrorComment)
                                End If
                            Next j

                        Case 2
                            Call WriteLog(False, "Patten2 点灯")
                            For j = 1 To 3

                                rcv = TSSendRecv_NonDisp(mCommand_LED_ON(j * 2 - 1, 1), 3)

                                If rcv <> mCommand_LED_ON(j * 2 - 1, 2) Then
                                    Call WriteLog(False, "RCV" & i & ":" & rcv)
                                    mErrorComment = mCommand_LED_ON(j * 2 - 1, 0) & "RCV Data 異常"
                                    Throw New ApplicationException(mErrorComment)
                                End If
                            Next j
                        Case 3
                            Call WriteLog(False, "全消灯")
                            For j = 1 To 4

                                rcv = TSSendRecv_NonDisp(mCommand_LED_OFF(j, 1), 3)

                                If rcv <> mCommand_LED_OFF(j, 2) Then
                                    Call WriteLog(False, "RCV" & i & ":" & rcv)
                                    mErrorComment = mCommand_LED_OFF(j, 0) & "RCV Data 異常"
                                    Throw New ApplicationException(mErrorComment)
                                End If
                            Next j

                            Call WriteLog(False, "Patten3 点灯")
                            For j = 1 To 3

                                rcv = TSSendRecv_NonDisp(mCommand_LED_ON(j * 2, 1), 3)

                                If rcv <> mCommand_LED_ON(j * 2, 2) Then
                                    Call WriteLog(False, "RCV" & i & ":" & rcv)
                                    mErrorComment = mCommand_LED_ON(j * 2 - 1, 0) & "RCV Data 異常"
                                    Throw New ApplicationException(mErrorComment)
                                End If
                            Next j

                    End Select

                    Test_Count += 1

                    Call cEI.GAZO_TEST()

                    Try
                        Select Case i
                            Case 1 : If cEI.RESULT <> "1,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                            Case 2 : If cEI.RESULT <> "0,1,0,1,0,1,0,1,1,1,1,1" Then GoTo ErrCheck
                            Case 3 : If cEI.RESULT <> "0,0,1,0,1,0,1,0,1,1,1,1" Then GoTo ErrCheck

                        End Select

                        Call WriteLog(False, "パターン" & i & "検査OK")

                        GoTo NextGo

ErrCheck:
                        NG_Count += 1
                        Call WriteLog(False, "目視検査")

                        'メッセージ表示
                        dfrm = New frm500_MessagePopUp
                        dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7x\" & "LED" & i & ".ppt"
                        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                        Call dfrm.frmShow()

                        Do
                            'wait push button -------------------
                            If dfrm.DIALOG_RESULT = "NG" Then
                                dfrm.Close()
                                dfrm = Nothing
                                mErrorComment = FuncName & "検査 異常 "
                                mTestSpec = FuncName   'Rev0.20add
                                Throw New ApplicationException(mErrorComment)
                            End If

                            If dfrm.DIALOG_RESULT = "OK" Then
                                Call WriteLog(False, "パターン" & i & "検査OK")
                                Exit Do
                            End If

                            DoEvents()

                        Loop

                        Call WriteLog(False, "Patten " & i & " OK")
                        Call WriteLog(False, "Result : " & cEI.RESULT)

                        dfrm.Close()
                        dfrm = Nothing

                    Catch ex As Exception

                        Call WriteLog(False, mErrorComment)
                        'Throw New ApplicationException(mErrorComment)
                    Finally
                        Call TSSendRecv_NonDisp("TS,7,0,0,FFFE3A02,2,0FFF,", 3)
                        Call TSSendRecv_NonDisp("TS,7,0,0,3880B22A,1,0,", 3)
                        Call TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,0,", 3)
                        Call TSSendRecv_NonDisp("TS,7,0,0,3880B229,1,0,", 3)

                        Call TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,C0,", 3)


                    End Try

NextGo:

                Next

                '------------------------------------- Ver 0.27 JDI ADD End --------------------------------------------

            Case False      'カメラなしのとき
                '------------------------- Ver 0.27 JDI ADD Start -----------------------
                Try
                    For i = 0 To 3
                        Test_Count += 1                                 'Ver 0.24 JDI ADD
                        NG_Count += 1                                   'Ver 0.24 JDI ADD
                        Call WriteLog(False, "目視検査")                'Ver 0.24 JDI ADD

                        Select Case i
                            Case 0

                            Case 1
                                Call WriteLog(False, "全消灯")
                                For j = 0 To 4

                                    rcv = TSSendRecv_NonDisp(mCommand_LED_OFF(j, 1), 3)

                                    If rcv <> mCommand_LED_OFF(j, 2) Then
                                        Call WriteLog(False, "RCV" & i & ":" & rcv)
                                        mErrorComment = mCommand_LED_OFF(j, 0) & "RCV Data 異常"
                                        Throw New ApplicationException(mErrorComment)
                                    End If
                                Next j

                            Case 2
                                Call WriteLog(False, "Patten2 点灯")
                                For j = 1 To 3

                                    rcv = TSSendRecv_NonDisp(mCommand_LED_ON(j * 2 - 1, 1), 3)

                                    If rcv <> mCommand_LED_ON(j * 2 - 1, 2) Then
                                        Call WriteLog(False, "RCV" & i & ":" & rcv)
                                        mErrorComment = mCommand_LED_ON(j * 2 - 1, 0) & "RCV Data 異常"
                                        Throw New ApplicationException(mErrorComment)
                                    End If
                                Next j
                            Case 3
                                Call WriteLog(False, "全消灯")
                                For j = 1 To 4

                                    rcv = TSSendRecv_NonDisp(mCommand_LED_OFF(j, 1), 3)

                                    If rcv <> mCommand_LED_OFF(j, 2) Then
                                        Call WriteLog(False, "RCV" & i & ":" & rcv)
                                        mErrorComment = mCommand_LED_OFF(j, 0) & "RCV Data 異常"
                                        Throw New ApplicationException(mErrorComment)
                                    End If
                                Next j

                                Call WriteLog(False, "Patten3 点灯")

                                For j = 1 To 3

                                    rcv = TSSendRecv_NonDisp(mCommand_LED_ON(j * 2, 1), 3)

                                    If rcv <> mCommand_LED_ON(j * 2, 2) Then
                                        Call WriteLog(False, "RCV" & i & ":" & rcv)
                                        mErrorComment = mCommand_LED_ON(j * 2, 0) & "RCV Data 異常"
                                        Throw New ApplicationException(mErrorComment)
                                    End If

                                Next j
                        End Select

                        ''メッセージ表示
                        dfrm = New frm500_MessagePopUp

                        If i = 0 Then
                            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7x\" & Mid(Model, 3, 4) & "_Filter.ppt"
                        Else
                            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7x\" & "LED" & i & ".ppt"
                        End If

                        dfrm.DISP_TITLENAME = "Filter & LED 검사"
                        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                        Call dfrm.frmShow()

                        Do
                            'wait push button -------------------
                            If dfrm.DIALOG_RESULT = "NG" Then
                                mErrorComment = "LED 異常"
                                mTestSpec = "Patten" & i + 1    'Rev0.20add
                                Call WriteLog(False, mErrorComment)

                                dfrm.Close()
                                dfrm = Nothing
                                Throw New ApplicationException(mErrorComment)
                            End If

                            If dfrm.DIALOG_RESULT = "OK" Then
                                If i = 0 Then
                                    Call WriteLog(False, "Filter検査OK")
                                Else
                                    Call WriteLog(False, "パターン" & i & "検査OK")
                                End If

                                dfrm.Close()
                                dfrm = Nothing
                                Exit Do
                            End If

                            DoEvents()
                        Loop
                    Next i
                Catch ex As Exception

                Finally
                    Call TSSendRecv_NonDisp("TS,7,0,0,FFFE3A02,2,0FFF,", 3)
                    Call TSSendRecv_NonDisp("TS,7,0,0,3880B22A,1,0,", 3)
                    Call TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,0,", 3)
                    Call TSSendRecv_NonDisp("TS,7,0,0,3880B229,1,0,", 3)

                    Call TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,C0,", 3)
                End Try

        End Select
        '-------------------------- Ver 0.27 JDI ADD End ------------------------

        '------------------------- Ver 0.27 JDI DEL Start -----------------------
        ''メッセージ表示
        'dfrm = New frm500_MessagePopUp
        'dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\LED.ppt"
        ''dfrm.DISP_TITLENAME = "LED 検査"           'Rev0.11DEL
        'dfrm.DISP_TITLENAME = "LED 검사"           'Rev0.11ADD
        'dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        'Call dfrm.frmShow()
        'Do
        'For i = 0 To 17
        '    Test_Count += 1                                 'Ver 0.24 JDI ADD
        '    NG_Count += 1                                   'Ver 0.24 JDI ADD
        '    Call WriteLog(False, "目視検査")                'Ver 0.24 JDI ADD

        '    Select Case i
        '        '-----< Rev0.11DEL >-----
        '        'Case 0 : FuncName = "全消灯" : SendCom = "TS,7,0,0,3880B21B,1,0," : RecvCom = "TS,7,0,0,0,1,3880B21B,0,"
        '        'Case 1 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,3880B21B,1,C0," : RecvCom = "TS,7,0,0,1,1,3880B21B,0,"
        '        'Case 2 : FuncName = "RDY,RUN点灯" : SendCom = "TS,7,0,0,3880B228,1,80," : RecvCom = "TS,7,0,0,0,1,3880B228,80,"
        '        'Case 3 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,3880B228,1,0," : RecvCom = "TS,7,0,0,0,1,3880B228,0,"
        '        'Case 4 : FuncName = "全消灯" : SendCom = "TS,7,0,0,3880B21B,1,0," : RecvCom = "TS,7,0,0,0,1,3880B21B,0,"
        '        'Case 5 : FuncName = "ALM点灯" : SendCom = "TS,7,0,0,3880B229,1,80," : RecvCom = "TS,7,0,0,0,1,3880B229,80,"
        '        'Case 6 : FuncName = "全消灯" : SendCom = "TS,7,0,0,3880B229,1,0," : RecvCom = "TS,7,0,0,0,1,3880B229,0,"
        '        'Case 7 : FuncName = "ERR点灯" : SendCom = "TS,7,0,0,3880B22A,1,80," : RecvCom = "TS,7,0,0,0,1,3880B22A,80,"
        '        'Case 8 : FuncName = "全消灯" : SendCom = "TS,7,0,0,3880B22A,1,0," : RecvCom = "TS,7,0,0,0,1,3880B22A,0,"
        '        'Case 9 : FuncName = "SD点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FEF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FEF,"
        '        'Case 10 : FuncName = "全消灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        'Case 11 : FuncName = "EXE点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FDF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FDF,"
        '        'Case 12 : FuncName = "全消灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        'Case 13 : FuncName = "US1点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FBF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FBF,"
        '        'Case 14 : FuncName = "全消灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        'Case 15 : FuncName = "US2点灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0F7F," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,F7F,"
        '        'Case 16 : FuncName = "全消灯" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        'Case 17 : FuncName = "RDY点灯" : SendCom = "TS,7,0,0,3880B21B,1,C0," : RecvCom = "TS,7,0,0,1,1,3880B21B,0,"
        '        Case 0 : FuncName = "전소등" : SendCom = "TS,7,0,0,3880B21B,1,0," : RecvCom = "TS,7,0,0,0,1,3880B21B,0,"
        '        Case 1 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,3880B21B,1,C0," : RecvCom = "TS,7,0,0,1,1,3880B21B,0,"
        '        Case 2 : FuncName = "RDY,RUN점등" : SendCom = "TS,7,0,0,3880B228,1,80," : RecvCom = "TS,7,0,0,0,1,3880B228,80,"
        '        Case 3 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,3880B228,1,0," : RecvCom = "TS,7,0,0,0,1,3880B228,0,"
        '        Case 4 : FuncName = "전소등" : SendCom = "TS,7,0,0,3880B21B,1,0," : RecvCom = "TS,7,0,0,0,1,3880B21B,0,"
        '        Case 5 : FuncName = "ALM점등" : SendCom = "TS,7,0,0,3880B229,1,80," : RecvCom = "TS,7,0,0,0,1,3880B229,80,"
        '        Case 6 : FuncName = "전소등" : SendCom = "TS,7,0,0,3880B229,1,0," : RecvCom = "TS,7,0,0,0,1,3880B229,0,"
        '        Case 7 : FuncName = "ERR점등" : SendCom = "TS,7,0,0,3880B22A,1,80," : RecvCom = "TS,7,0,0,0,1,3880B22A,80,"
        '        Case 8 : FuncName = "전소등" : SendCom = "TS,7,0,0,3880B22A,1,0," : RecvCom = "TS,7,0,0,0,1,3880B22A,0,"
        '        Case 9 : FuncName = "SD점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FEF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FEF,"
        '        Case 10 : FuncName = "전소등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        Case 11 : FuncName = "EXE점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FDF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FDF,"
        '        Case 12 : FuncName = "전소등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        Case 13 : FuncName = "US1점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FBF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FBF,"
        '        Case 14 : FuncName = "전소등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        Case 15 : FuncName = "US2점등" : SendCom = "TS,7,0,0,FFFE3A02,2,0F7F," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,F7F,"
        '        Case 16 : FuncName = "전소등" : SendCom = "TS,7,0,0,FFFE3A02,2,0FFF," : RecvCom = "TS,7,0,0,0,2,FFFE3A02,FFF,"
        '        Case 17 : FuncName = "RDY점등" : SendCom = "TS,7,0,0,3880B21B,1,C0," : RecvCom = "TS,7,0,0,1,1,3880B21B,0,"
        '            '---------------------------------
        '    End Select

        '    Try
        '        rcv = TSSendRecv_NonDisp(SendCom, 3)

        '        Call mSecWait(500)

        '        '------------------------- Ver 0.25 JDI ADD Start -----------------------
        '        Call WriteLog(False, "RCV:" & rcv)

        '        If rcv <> RecvCom Then
        '            mErrorComment = "RCV Data 異常"

        '            Throw New ApplicationException(mErrorComment)
        '        End If
        '        '-------------------------- Ver 0.25 JDI ADD End ------------------------

        '    Catch ex As Exception
        '        dfrm.Close()
        '        dfrm = Nothing
        '        Throw New ApplicationException(mErrorComment)
        '    End Try
        '    'Call WriteLog(False, FuncName & " OK " & rcv)

        '    'wait push button -------------------
        '    If dfrm.DIALOG_RESULT = "NG" Then
        '        mErrorComment = "LED 異常"
        '        mTestSpec = FuncName    'Rev0.20add
        '        Call WriteLog(False, mErrorComment)

        '        dfrm.Close()
        '        dfrm = Nothing
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    If dfrm.DIALOG_RESULT = "OK" Then
        '        rcv = TSSendRecv_NonDisp("TS,7,0,0,FFFE3A02,2,0FFF,", 3)
        '        rcv = TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,C0,", 3)
        '        Exit Do
        '    End If

        'Next
        'Loop
        '-------------------------- Ver 0.27 JDI DEL End ------------------------

        Call WriteLog(False, "検査数：" & Test_Count)                                                    'Ver 0.24 JDI ADD
        Call WriteLog(False, "目視検査数：" & NG_Count)                                                  'Ver 0.24 JDI ADD
        Call WriteLog(False, "自動検査合格率：" & Format((Test_Count - NG_Count) / Test_Count * 100, "0.00") & "%")     'Ver 0.24 JDI ADD

        '全点灯確認
        Try
            '------------------------- Ver 0.25 JDI DEL Start -----------------------
            'SendCom = "TS,7,0,0,FFFE3A02,2,0000,"
            'rcv = TSSendRecv_NonDisp(SendCom, 3)
            'SendCom = "TS,7,0,0,3880B21B,1,C0,"
            'rcv = TSSendRecv_NonDisp(SendCom, 3)
            'SendCom = "TS,7,0,0,3880B228,1,80,"
            'rcv = TSSendRecv_NonDisp(SendCom, 3)
            'SendCom = "TS,7,0,0,3880B229,1,80,"
            'rcv = TSSendRecv_NonDisp(SendCom, 3)
            'SendCom = "TS,7,0,0,3880B22A,1,80,"
            'rcv = TSSendRecv_NonDisp(SendCom, 3)

            'Call mSecWait(500)
            '-------------------------- Ver 0.25 JDI DEL End ------------------------

            '------------------------- Ver 0.25 JDI ADD Start -----------------------
            For i = 1 To 5
                Select Case i
                    Case 1 : SendCom = "TS,7,0,0,FFFE3A02,2,0000," : RecvCom = "TS,7,0,0,1,2,FFFE3A02,F00,"
                    Case 2 : SendCom = "TS,7,0,0,3880B21B,1,C0," : RecvCom = "TS,7,0,0,1,1,3880B21B,0,"
                    Case 3 : SendCom = "TS,7,0,0,3880B228,1,80," : RecvCom = "TS,7,0,0,0,1,3880B228,80,"
                    Case 4 : SendCom = "TS,7,0,0,3880B229,1,80," : RecvCom = "TS,7,0,0,0,1,3880B229,80,"
                    Case 5 : SendCom = "TS,7,0,0,3880B22A,1,80," : RecvCom = "TS,7,0,0,0,1,3880B22A,80,"

                End Select
                rcv = TSSendRecv_NonDisp(SendCom, 3)

                'Call mSecWait(500)

                If rcv <> RecvCom Then
                    Call WriteLog(False, "RCV" & i & ":" & rcv)
                    mErrorComment = "RCV Data 異常"
                    Throw New ApplicationException(mErrorComment)
                End If
            Next i

            Call mSecWait(500)
            '-------------------------- Ver 0.25 JDI ADD End ------------------------


            If fCam = True Then
                Call cEI.GAZO_TEST()
                Call WriteLog(False, "全点灯")
                Call WriteLog(False, "目視検査のみ")     'Ver 0.24 JDI ADD
                Call WriteLog(False, cEI.RESULT)
                'Call WriteLog(False, cEI.ResultData)
            End If

            'メッセージ表示
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\ALLLED.ppt"
            'dfrm.DISP_TITLENAME = "LED 検査"           'Rev0.11DEL
            dfrm.DISP_TITLENAME = "LED 검사"           'Rev0.11ADD
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "LED全点灯検査 異常 "
                    mTestSpec = "LED全点灯"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

            Loop

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call TSSendRecv_NonDisp("TS,7,0,0,FFFE3A02,2,0FFF,", 3)
            Call TSSendRecv_NonDisp("TS,7,0,0,3880B22A,1,0,", 3)
            Call TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,0,", 3)
            Call TSSendRecv_NonDisp("TS,7,0,0,3880B229,1,0,", 3)

            Call TSSendRecv_NonDisp("TS,7,0,0,3880B21B,1,C0,", 3)
            dfrm.Close()
            dfrm = Nothing
        End Try

        If fCam = True Then cEI = Nothing

        Call WriteLog(False, "LED検査 正常終了")

    End Sub

    Public Sub I27(ByVal iName As String)
        '==============================================================================================================
        ' Function : I27:ﾛｰﾀﾘｰｽｲｯﾁLED検査
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
        Call WriteLog(False, "SW検査中・・・")


        'メッセージ表示
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\SW.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "ロータリー・スイッチ検査 異常 "
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "ロータリー・スイッチ検査 正常")

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I28(ByVal iName As String)
        '==============================================================================================================
        ' Function : I28:SDｶｰﾄﾞ検査-ｶｰﾄﾞ電源/過電流保護
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
        Dim ADDevice As String
        Dim YDDevice As String
        Dim RecvDataLE As String = ""
        'Dim Result As String
        'Dim sValue As String
        Dim dValue As Single
        Dim StdMax As Single
        Dim StdMin As Single
        Dim SigName As String
        Dim TestStd_SDDWN_H As Single = 0.2
        Dim TestStd_SDDWN_L As Single = 0
        Dim TestStd_SDPWR_H As Single
        Dim TestStd_SDPWR_L As Single
        Dim mV3P3 As Double
        Dim cMeas As New cls140_VoltageMeasure

        Try
            'ﾘﾚｰ設定（GND）
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U1
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "509" & ",001,1", RecvDataLE)   'GND側設定
            'If cLnk.ERR_COMMENT <> "" Then
            '    mErrorComment = cLnk.ERR_COMMENT
            '    Throw New ApplicationException(mErrorComment)
            'End If

            'VCC3P3測定
            'ADDevice = "0002"     'Ver 0.29 JDI DEL
            'ﾘﾚｰ設定ON
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "5" & Strings.Right(ADDevice, 2) & ",001,1", RecvDataLE)
            'If cLnk.ERR_COMMENT <> "" Then
            '    mErrorComment = cLnk.ERR_COMMENT
            '    Throw New ApplicationException(mErrorComment)
            'End If
            'DMM測定
            'Call WaitStart(1)
            'mV3P3 = DMM_7555_Measure()

            'v0.07　DMM測定
            'mV3P3 = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))      'Ver 0.29 JDI DEL

            'TestStd_SDPWR_H = CSng(mV3P3 + 0.005)                      'Ver 0.29 JDI DEL
            'TestStd_SDPWR_L = CSng(mV3P3 - 0.05)                       'Ver 0.29 JDI DEL

            'ﾘﾚｰ設定OFF
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "5" & Strings.Right(ADDevice, 2) & ",001,0", RecvDataLE)
            'Call WaitStart(1)

            '-------------------- Ver 0.29 JDI ADD Start --------------------
            Dim dbName As String = ""
            Dim tbName As String = ""
            Dim sSQL As String = ""
            Dim rsData() As String = {""}
            Dim rdData() As String = {""}
            Dim rfName() As String = {""}
            Dim rdName() As String = {""}

            Dim DataCount As Integer = 0

            dbName = mAppPath & "210 Result Data\Temporary.mdb"
            tbName = "tblResult"
            sSQL = "SELECT StartDateTime,MSCODE,SerialNo FROM " & tbName

            Try
                cADO = New cls010_MDB_Operation
                cADO.DB_PATH = dbName
                cADO.TB_NAME = tbName
                cADO.DB_SQL = sSQL

                Call cADO.DB_READ_JET(rfName, rsData)

                DataCount = cADO.DB_DATA_COUNT

            Catch ex As Exception
                mErrorComment = cADO.ERR_COMMENT
                If mErrorComment = "" Then mErrorComment = ex.ToString
                Throw New ApplicationException(mErrorComment)
            Finally
                cADO = Nothing
            End Try

            If DataCount = 0 Then
                ADDevice = "0002"
                
                mV3P3 = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))
                Call WriteLog(True, "DMM V3P3 : " & mV3P3 & "V")

                TestStd_SDPWR_H = CSng(mV3P3 + 0.005)
                TestStd_SDPWR_L = CSng(mV3P3 - 0.05)
            Else
                dbName = mAppPath & "210 Result Data\F3SP7XData.mdb"
                tbName = "tblResult"

                'データがあるかどうかチェック
                sSQL = "SELECT * FROM " & tbName & " "
                sSQL &= "WHERE 検査日時='" & rsData(0) & "' AND MSCODE='" & rsData(1) & "' AND SerialNo='" & rsData(2) & "'"
                Try
                    cADO = New cls010_MDB_Operation
                    cADO.DB_PATH = dbName
                    cADO.TB_NAME = tbName
                    cADO.DB_SQL = sSQL
                    Call cADO.DB_READ_JET(rdName, rdData)
                    DataCount = cADO.DB_DATA_COUNT
                Catch ex As Exception
                    mErrorComment = cADO.ERR_COMMENT
                    If mErrorComment = "" Then mErrorComment = ex.ToString
                    Throw New ApplicationException(mErrorComment)
                Finally
                    cADO = Nothing
                End Try

                If DataCount = 0 Then
                    ADDevice = "0002"

                    mV3P3 = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))
                    Call WriteLog(True, "DMM V3P3 : " & mV3P3 & "V")

                    TestStd_SDPWR_H = CSng(mV3P3 + 0.005)
                    TestStd_SDPWR_L = CSng(mV3P3 - 0.05)
                Else
                    Call WriteLog(True, "DB V3P3 : " & rdData(4) & "V")
                    TestStd_SDPWR_H = CSng(CDbl(rdData(4)) + 0.005)
                    TestStd_SDPWR_L = CSng(CDbl(rdData(4)) - 0.05)
                End If

            End If

            '--------------------- Ver 0.29 JDI ADD End ---------------------


            'SD電源 -------------------------------------
            SigName = "SDPWR"
            ADDevice = "0007"
            StdMax = TestStd_SDPWR_H : StdMin = TestStd_SDPWR_L
            Call WriteLog(True, SigName & "検査中・・・")
            ''ﾘﾚｰ設定ON
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "5" & Strings.Right(ADDevice, 2) & ",001,1", RecvDataLE)
            'If cLnk.ERR_COMMENT <> "" Then
            '    mErrorComment = cLnk.ERR_COMMENT
            '    Throw New ApplicationException(mErrorComment)
            'End If
            ''DMM測定
            'Call WaitStart(1)
            'dValue = DMM_7555_Measure()

            'v0.07　DMM測定
            dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

            Call WriteLog(True, "電源ON時" & SigName & "測定中・・・")

            If dValue > StdMax Or dValue < StdMin Then
                mErrorComment = "電源ON時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
                mTestMeasure = dValue.ToString     'Rev0.20add
                mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
                Call WriteLog(True, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(True, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString)

        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        Finally
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "501" & ",016,0000000000000000", RecvDataLE)
            'cLnk = Nothing
        End Try

        'ResultData.SDPower = dValue

        '電源ショート -------------------------------------
        YDDevice = "903"
        Call cLnk.F3PCLinkCommunication("01" & "BWRY" + YDDevice + ",001,1", RecvDataLE)
        RecvDataLE = Strings.Mid(RecvDataLE, 3, 2)
        If RecvDataLE <> "OK" Then
            mErrorComment = "SD電源ｼｮｰﾄ 異常 "
            mTestSpec = "SD電源ｼｮｰﾄ"    'Rev0.20add
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "SD電源ｼｮｰﾄ 正常")

        Call WaitStart(1)

        YDDevice = "903"
        Call cLnk.F3PCLinkCommunication("01" & "BWRY" + YDDevice + ",001,0", RecvDataLE)
        RecvDataLE = Strings.Mid(RecvDataLE, 3, 2)
        If RecvDataLE <> "OK" Then
            mErrorComment = "SD電源ｼｮｰﾄ解除 異常 "
            mTestSpec = "SD電源ｼｮｰﾄ解除"    'Rev0.20add
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "SD電源ｼｮｰﾄ解除 正常")
        Call mSecWait(500)

        'SD電源 -------------------------------------
        SigName = "SDDWN"
        ADDevice = "0007"
        StdMax = TestStd_SDDWN_H : StdMin = TestStd_SDDWN_L

        Call WriteLog(False, SigName & "検査中・・・")

        'Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
        'RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
        'Result = Strings.Left(RecvDataLE, 2)
        'If Result <> "OK" Then
        '    mErrorComment = SigName & "測定返信 異常 " & RecvDataLE
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End If
        'sValue = Mid(RecvDataLE, 3, 4)
        'dValue = CSng(CDec("&h" & sValue) / (20000 / 10))

        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))
        cMeas = Nothing

        If dValue < 0 Then dValue = Math.Abs(dValue)

        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
            mTestMeasure = dValue.ToString     'Rev0.20add
            mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString("0.00"))

        'メッセージ表示 -------------------------------------------------
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\SDChange.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "ｼｮｰﾄ後のSD認識　異常"
                    mTestSpec = "ｼｮｰﾄ後のSD認識"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "ｼｮｰﾄ後のSD認識　正常")

        ' ---------------------------------------------------------------------------------------
    End Sub

    Public Sub I29(ByVal iName As String)
        '==============================================================================================================
        ' Function : I29:SDｶｰﾄﾞ検査-ﾗｲﾄﾌﾟﾛﾃｸﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mMeas As String = ""
        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test   
        ' ---------------------------------------------------------------------------------------
        Dim rcv As String

        'ﾃｽﾄﾓｰﾄﾞ移行
        'Call I05(iName)     ' Ver 0.29 JDI DEL

        For i = 1 To 1
            Select Case i
                'Case 0 : FuncName = "準備：入力ﾎﾟｰﾄの設定" : SendCom = "TS,7,0,0,FFFE398E,2,A0AA," : RecvCom = "TS,7,0,0,0,2,FFFE398E,A0AA,"
                Case 1 : FuncName = "WriteProtect有り確認" : SendCom = "TS,6,0,0,FFFE3982,2," : RecvCom = "TS,6,0,0,0,2,10,FFFE3982,"
            End Select

            Call WriteLog(False, FuncName & "中・・・")
            rcv = TSSendRecv(SendCom, 3)
            If rcv <> RecvCom Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = RecvCom    'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Call I06(iName)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, FuncName & " OK ")

            'Call WaitTimer(1)   ' Ver 0.29 JDI DEL
        Next

        ' ---------------------------------------------------------------------------------------
        'ﾃｽﾄﾓｰﾄﾞ終了
        'Call I06(iName)    ' Ver 0.29 JDI DEL

    End Sub

    Public Sub I30(ByVal iName As String)
        '==============================================================================================================
        ' Function : I30:ﾗﾀﾞｰ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mMeas As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Dim FunctionName As String = "ﾗﾀﾞｰ検査"

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, FunctionName & "中・・・")

        Dim IDevice As String = ""
        Dim RecvDataLE As String = ""
        Dim Result As String
        Dim RunWait As Integer

        Call WriteLog(False, FunctionName & "-SDｱｸｾｽ/SET SW検査")

        'メッセージ表示 -------------------------------------------------
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7x\SDStart.ppt"
            dfrm.DISP_TITLENAME = FunctionName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "ラダー運転で中止が押されました。"
                    mTestSpec = "ラダーStart"   'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "ﾗﾀﾞｰ運転開始")

        RunWait = 5
        Call WriteLog(False, "ﾗﾀﾞｰ運転 " & RunWait & "秒待ち")

        Call WaitStart(RunWait)

        'Iレジスタ読み込み -------------------------------------

        Call WriteLog(False, "Iレジスタ読み込み")

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_SP

            For i = 0 To 1
                If i = 0 Then IDevice = "I00012"
                If i = 1 Then IDevice = "I00013"
                Call WriteLog(False, "Send : " & "01" & "BRD" & IDevice + ",001")
                Call cLnk.F3PCLinkCommunication("01" & "BRD" & IDevice + ",001", RecvDataLE)
                Call WriteLog(False, "Recv : " & RecvDataLE)

                Result = Strings.Mid(RecvDataLE, 3, 2)
                If Result <> "OK" Then
                    mErrorComment = "Iレジスタ受信　異常 " & RecvDataLE
                    mTestSpec = "OK"   'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                If Mid(RecvDataLE, 5, 1) <> CStr(i) Then
                    mErrorComment = FunctionName & "　異常 " & Mid(RecvDataLE, 3, 4) & " (正常は000" & CStr(i) & ")"
                    Throw New ApplicationException(mErrorComment)
                End If
            Next

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, FunctionName & "　正常 " & Mid(RecvDataLE, 3, 4))

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I31(ByVal iName As String)
        '==============================================================================================================
        ' Function : I31:EthernetLED検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 10

        Dim rCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "EthernetLED検査中・・・")

        Dim rcv As String

        'メッセージ表示
        Try
            'LANｱﾀﾞﾌﾟﾀdisable
            cPCSP = New cls060_PCNetSPEED
            Call cPCSP.DisableNetAdapter()

            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\LAN LED消灯確認.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "EthernetLED消灯検査 異常 "
                    mTestSpec = "EthernetLED消灯"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()
                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            dfrm.Close()
            dfrm = Nothing
            Call cPCSP.EnableNetAdapter()
            cPCSP = Nothing

        End Try

        Call WaitStart(7)

        'メッセージ表示 COM・LED点滅確認
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\LAN LED点滅確認.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()
            'COM点滅
            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "EthernetLED検査 異常 "
                    mTestSpec = "EthernetLED点滅"   'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Call WriteLog(False, "EthernetLED検査 正常")
                    Exit Do
                End If

                rcv = TSSendRecv_NonDisp("TS,13,0,0,", 5)
                DoEvents()
                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            dfrm.Close()
            dfrm = Nothing
        End Try

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I32(ByVal iName As String)
        '==============================================================================================================
        ' Function : I32:ﾃﾞｰﾀ書込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mReadLine As String = ""
        Dim mLOG As String = ""                       'v0.01

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "ﾃﾞｰﾀ書込中・・・")

        Dim WriteMACAdres As String
        Dim WriteSerialAdres As String
        Dim HardRev As String = ""

        Dim FuncName As String
        Dim sMac(2) As String
        Dim sSer(4) As String
        Dim AscSer As String
        Dim sDate(4) As String
        Dim AscDay As String
        Dim sDay As String
        Dim sHRev As String
        Dim i As Integer
        Dim rcv As String
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Dim sMacR(2) As String
        Dim sMacAll As String
        Dim sSerR(4) As String
        Dim sSerAll As String
        Dim chSer As String
        Dim sDayR(4) As String
        Dim sDayAll As String
        Dim chDay As String
        Dim sHRevR As String
        'Dim chRev As String

        'データ作成 ---------------------------------------------------
        Call WriteLog(False, "ﾃﾞｰﾀ作成")

        'SER
        AscSer = ""
        'WriteSerialAdres = IA0()
        WriteSerialAdres = SerialRead()
        For i = 1 To Len(WriteSerialAdres)
            AscSer = AscSer & Hex(Asc(Mid(WriteSerialAdres, i, 1)))
        Next
        sSer(0) = Mid(AscSer, 1, 4)
        sSer(1) = Mid(AscSer, 5, 4)
        sSer(2) = Mid(AscSer, 9, 4)
        sSer(3) = Mid(AscSer, 13, 4)
        sSer(4) = Mid(AscSer, 17, 2) & "00"
        'Call WriteLog(False, "ｼﾘｱﾙNo.：" & sSer(0) & sSer(1) & sSer(2) & sSer(3) & sSer(4))    'v0.01
        mLOG = "ｼﾘｱﾙNo.：" & sSer(0) & sSer(1) & sSer(2) & sSer(3) & sSer(4) & vbCrLf           'v0.01

        'MAC
        WriteMACAdres = IA1()

        sMac(0) = Mid(WriteMACAdres, 1, 4)
        sMac(1) = Mid(WriteMACAdres, 5, 4)
        sMac(2) = Mid(WriteMACAdres, 9, 4)
        'Call WriteLog(False, "MACｱﾄﾞﾚｽ：" & sMac(0) & sMac(1) & sMac(2))                       'v0.01
        mLOG &= "MACｱﾄﾞﾚｽ：" & sMac(0) & sMac(1) & sMac(2) & vbCrLf                             'v0.01

        'Day
        AscDay = ""
        sDay = Format(DateTime.Now, "yyyy/MM/dd")
        For i = 1 To Len(sDay)
            AscDay = AscDay & Hex(Asc(Mid(sDay, i, 1)))
        Next
        sDate(0) = Mid(AscDay, 1, 4)
        sDate(1) = Mid(AscDay, 5, 4)
        sDate(2) = Mid(AscDay, 9, 4)
        sDate(3) = Mid(AscDay, 13, 4)
        sDate(4) = Mid(AscDay, 17, 4)
        'Call WriteLog(False, "日付：" & sDate(0) & sDate(1) & sDate(2) & sDate(3) & sDate(4))  'v0.01
        mLOG &= "日付：" & sDate(0) & sDate(1) & sDate(2) & sDate(3) & sDate(4) & vbCrLf        'v0.01

        'HardRev
        'zeroForm = ""
        HardRev = HardRevRead()
        sHRev = "00" & HardRev
        'sHRev = "5050"          'PP用
        'Call WriteLog(False, "HardRev：" & sHRev)      'v0.01
        mLOG &= "HardRev：" & sHRev                     'v0.01
        Call WriteLog(False, mLOG) : mLOG = ""          'v0.01

        '書き込み開始 --------------------------------------------------
        FuncName = "特定領域ﾌﾗｯｼｭ消去"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = TSSendRecv("TS,22,0,0,203FC000,203FC03F,1,", 3)
        If rcv <> "TS,22,0,0,0,203FC000," Then
            mErrorComment = FuncName & "　異常 " & rcv
            mTestSpec = "TS,22,0,0,0,203FC000,"    'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK ")

        ' ------------------------------------------------------
        'SER書き込み
        Call WriteLog(False, "SER書き込み中・・・")
        For i = 0 To 7
            FuncName = "SER書き込み-" & CInt(i + 1) & "/5"
            Select Case i
                Case 0 : SendCom = "TS,23,0,0,203FC000," & sSer(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC000,"
                Case 1 : SendCom = "TS,23,0,0,203FC002," & sSer(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC002,"
                Case 2 : SendCom = "TS,23,0,0,203FC004," & sSer(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC004,"
                Case 3 : SendCom = "TS,23,0,0,203FC006," & sSer(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC006,"
                Case 4 : SendCom = "TS,23,0,0,203FC008," & sSer(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC008,"
                Case 5 : SendCom = "TS,23,0,0,203FC00A,0000,1," : RecvCom = "TS,23,0,0,1,203FC00A,"     '終端0x00書き込み
                Case 6 : SendCom = "TS,23,0,0,203FC00C,0000,1," : RecvCom = "TS,23,0,0,1,203FC00C,"     '終端0x00書き込み
                Case 7 : SendCom = "TS,23,0,0,203FC00E,0000,1," : RecvCom = "TS,23,0,0,1,203FC00E,"     '終端0x00書き込み
            End Select

            mLOG &= FuncName & vbCrLf                   'v0.01
            mLOG &= "Send : " & SendCom & vbCrLf        'v0.01
            rcv = TSSendRecv_NonDisp(SendCom, 3)
            mLOG &= "Recv : " & rcv & vbCrLf            'v0.01
            'If rcv <> RecvCom Then
            '    mErrorComment = FuncName & "　異常 " & rcv
            '    Throw New ApplicationException(mErrorComment)
            'End If
            'Call WriteLog(False, FuncName & " OK ")
        Next

        Call WriteLog(False, mLOG) : mLOG = ""          'v0.01
        ' ------------------------------------------------------
        'MACｱﾄﾞﾚｽ書き込み
        Call WriteLog(False, "MACｱﾄﾞﾚｽ書き込み中・・・")
        For i = 0 To 2
            FuncName = "MAC書き込み-" & CInt(i + 1) & "/3"
            Select Case i
                Case 0 : SendCom = "TS,23,0,0,203FC010," & sMac(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC010,"
                Case 1 : SendCom = "TS,23,0,0,203FC012," & sMac(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC012,"
                Case 2 : SendCom = "TS,23,0,0,203FC014," & sMac(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC014,"
            End Select

            mLOG &= FuncName & vbCrLf                   'v0.01
            mLOG &= "Send : " & SendCom & vbCrLf        'v0.01
            rcv = TSSendRecv_NonDisp(SendCom, 3)
            mLOG &= "Recv : " & rcv & vbCrLf            'v0.01
            'If rcv <> RecvCom Then
            '    mErrorComment = FuncName & "　異常 " & rcv
            '    Throw New ApplicationException(mErrorComment)
            'End If
            'Call WriteLog(False, FuncName & " OK ")
        Next

        Call WriteLog(False, mLOG) : mLOG = ""          'v0.01
        ' ------------------------------------------------------
        'Date書き込み
        Call WriteLog(False, "Date書き込み中・・・")
        For i = 0 To 7
            FuncName = "Date書き込み-" & CInt(i + 1) & "/5"
            Select Case i
                Case 0 : SendCom = "TS,23,0,0,203FC020," & sDate(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC020,"
                Case 1 : SendCom = "TS,23,0,0,203FC022," & sDate(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC022,"
                Case 2 : SendCom = "TS,23,0,0,203FC024," & sDate(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC024,"
                Case 3 : SendCom = "TS,23,0,0,203FC026," & sDate(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC026,"
                Case 4 : SendCom = "TS,23,0,0,203FC028," & sDate(i) & ",1," : RecvCom = "TS,23,0,0,1,203FC028,"
                Case 5 : SendCom = "TS,23,0,0,203FC02A,0000,1," : RecvCom = "TS,23,0,0,1,203FC02A,"
                Case 6 : SendCom = "TS,23,0,0,203FC02C,0000,1," : RecvCom = "TS,23,0,0,1,203FC02C,"
                Case 7 : SendCom = "TS,23,0,0,203FC02E,0000,1," : RecvCom = "TS,23,0,0,1,203FC02E,"
            End Select

            mLOG &= FuncName & vbCrLf                   'v0.01
            mLOG &= "Send : " & SendCom & vbCrLf        'v0.01
            rcv = TSSendRecv_NonDisp(SendCom, 3)
            mLOG &= "Recv : " & rcv & vbCrLf            'v0.01
            'If rcv <> RecvCom Then
            '    mErrorComment = FuncName & "　異常 " & rcv
            '    Throw New ApplicationException(mErrorComment)
            'End If
            'Call WriteLog(False, FuncName & " OK ")
        Next
        Call WriteLog(False, mLOG) : mLOG = ""          'v0.01
        ' ------------------------------------------------------
        FuncName = "H-Rev書き込み"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sHRev)

        rcv = TSSendRecv_NonDisp("TS,23,0,0,203FC030," & sHRev & ",1,", 3)
        If rcv <> "TS,23,0,0,0,203FC030," Then
            mErrorComment = FuncName & "　異常 " & rcv
            mTestSpec = "TS,23,0,0,0,203FC030,"    'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK ")
        Call WriteLog(False, "")

        '読み返し -----------------------------------------------------------------
        Call WriteLog(False, "読み返し確認")
        Call WriteLog(False, "")

        Dim strArry() As String
        ' ------------------------------------------------------
        'MAC読み込み
        For i = 0 To 2
            FuncName = "MAC読み込み-" & CInt(i + 1) & "/3"
            Select Case i
                Case 0 : SendCom = "TS,6,0,0,203FC010,2,"
                Case 1 : SendCom = "TS,6,0,0,203FC012,2,"
                Case 2 : SendCom = "TS,6,0,0,203FC014,2,"
            End Select

            'v0.01---------------------------------------------------------
            'Call WriteLog(False, FuncName & "中・・・")                
            'Call WriteLog(False, "書き込みデータ:" & sMac(i))   
            mLOG &= FuncName & vbCrLf
            mLOG &= "書き込みデータ:" & sMac(i) & vbCrLf
            'rcv = TSSendRecv(SendCom, 3)               
            mLOG &= "Send : " & SendCom & vbCrLf
            rcv = TSSendRecv_NonDisp(SendCom, 3)
            mLOG &= "Recv : " & rcv & vbCrLf
            '--------------------------------------------------------------

            strArry = Split(rcv, ",")
            sMacR(i) = strArry(6)
            If sMacR(i).ToString.Length <> 4 Then
                For j As Integer = 1 To 4 - sMacR(i).ToString.Length
                    sMacR(i) = "0" & sMacR(i)
                Next
            End If

            If sMacR(i) <> sMac(i) Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = sMac(i)    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
            'Call WriteLog(False, FuncName & " OK ")    'v0.01
            mLOG &= FuncName & " OK " & vbCrLf          'v0.01
        Next
        Call WriteLog(False, mLOG) : mLOG = ""          'v0.01

        '書き込み判定
        sMacAll = sMacR(0) & sMacR(1) & sMacR(2)
        If WriteMACAdres <> sMacAll Then
            mErrorComment = "MACｱﾄﾞﾚｽ読み返し　異常 " & sMacAll
            mTestSpec = WriteSerialAdres   'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "MACｱﾄﾞﾚｽ読み返し正常：" & sMacAll)
        Call WriteLog(False, "")

        ' ------------------------------------------------------
        'SER読み込み
        For i = 0 To 4
            FuncName = "SER読み込み-" & CInt(i + 1) & "/5"
            Select Case i
                Case 0 : SendCom = "TS,6,0,0,203FC000,2,"
                Case 1 : SendCom = "TS,6,0,0,203FC002,2,"
                Case 2 : SendCom = "TS,6,0,0,203FC004,2,"
                Case 3 : SendCom = "TS,6,0,0,203FC006,2,"
                Case 4 : SendCom = "TS,6,0,0,203FC008,2,"
            End Select

            'v0.01---------------------------------------------------------
            'Call WriteLog(False, FuncName & "中・・・約２秒待ち")
            'Call WriteLog(False, "書き込みデータ:" & sSer(i))
            mLOG &= FuncName & vbCrLf
            mLOG &= "書き込みデータ:" & sSer(i) & vbCrLf
            'rcv = TSSendRecv(SendCom, 3)
            mLOG &= "Send : " & SendCom & vbCrLf
            rcv = TSSendRecv_NonDisp(SendCom, 3)
            mLOG &= "Recv : " & rcv & vbCrLf
            '--------------------------------------------------------------

            strArry = Split(rcv, ",")
            sSerR(i) = strArry(6)
            If sSerR(i) <> sSer(i) Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = sSer(i)    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
            'Call WriteLog(False, FuncName & " OK ")
            mLOG &= FuncName & " OK " & vbCrLf          'v0.01
        Next
        Call WriteLog(False, mLOG) : mLOG = ""          'v0.01

        '書き込み判定
        sSerAll = sSerR(0) & sSerR(1) & sSerR(2) & sSerR(3) & sSerR(4)
        chSer = ""
        For i = 1 To Len(sSerAll) - 2 Step 2
            chSer = chSer & Chr(CInt("&H" & Mid(sSerAll, i, 2)))
        Next

        If WriteSerialAdres <> chSer Then
            mErrorComment = "ｼﾘｱﾙﾅﾝﾊﾞ読み返し　異常 " & chSer
            mTestSpec = WriteSerialAdres  'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "ｼﾘｱﾙﾅﾝﾊﾞ読み返し正常：" & chSer)
        Call WriteLog(False, "")

        ' ------------------------------------------------------
        'Date読み込み
        For i = 0 To 4
            FuncName = "Date読み込み-" & CInt(i + 1) & "/5"
            Select Case i
                Case 0 : SendCom = "TS,6,0,0,203FC020,2,"
                Case 1 : SendCom = "TS,6,0,0,203FC022,2,"
                Case 2 : SendCom = "TS,6,0,0,203FC024,2,"
                Case 3 : SendCom = "TS,6,0,0,203FC026,2,"
                Case 4 : SendCom = "TS,6,0,0,203FC028,2,"
            End Select

            'v0.01---------------------------------------------------------
            'Call WriteLog(False, FuncName & "中・・・約２秒待ち")
            'Call WriteLog(False, "書き込みデータ:" & sDate(i))
            mLOG &= FuncName & vbCrLf
            mLOG &= "書き込みデータ:" & sDate(i) & vbCrLf
            'rcv = TSSendRecv(SendCom, 3)
            mLOG &= "Send : " & SendCom & vbCrLf
            rcv = TSSendRecv_NonDisp(SendCom, 3)
            mLOG &= "Recv : " & rcv & vbCrLf
            '--------------------------------------------------------------

            strArry = Split(rcv, ",")
            sDayR(i) = strArry(6)
            If sDayR(i) <> sDate(i) Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = sDate(i)   'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
            'Call WriteLog(False, FuncName & " OK ")
            mLOG &= FuncName & " OK " & vbCrLf          'v0.01
        Next
        Call WriteLog(False, mLOG) : mLOG = ""          'v0.01

        '書き込み判定
        sDayAll = sDayR(0) & sDayR(1) & sDayR(2) & sDayR(3) & sDayR(4)
        chDay = ""
        For i = 1 To Len(sDayAll) Step 2
            chDay = chDay & Chr(CInt("&H" & Mid(sDayAll, i, 2)))
        Next

        If Format(DateTime.Now, "yyyy/MM/dd") <> chDay Then
            mErrorComment = "日付読み返し　異常 " & chDay
            mTestSpec = Format(DateTime.Now, "yyyy/MM/dd")   'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "日付読み返し正常：" & chDay)
        Call WriteLog(False, "")

        ' ------------------------------------------------------
        FuncName = "H-Rev読み込み"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = TSSendRecv("TS,6,0,0,203FC030,2,", 3)
        strArry = Split(rcv, ",")
        sHRevR = strArry(6)
        'sHRevR = Mid(rcv, 14, Len(rcv) - 23)

        If CInt("&H" & sHRevR) <> CInt("&H" & sHRev) Then         'sHRev
            mErrorComment = FuncName & "　異常 " & rcv
            mTestSpec = CInt("&H" & sHRev).ToString     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '書き込み判定
        'chRev = Strings.Left(sHRevR, 4)
        'chRev = CStr(Val("&H" & chRev))
        If CInt("&H" & sHRevR) <> CInt("&H" & sHRev) Then
            mErrorComment = "H-Rev読み返し　異常 " & "Rev=" & sHRev
            mTestSpec = sHRevR    'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "H-Rev読み返し正常：" & "Rev=" & sHRev)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I33(ByVal iName As String)
        '==============================================================================================================
        ' Function : I33:SDｶｰﾄﾞ検査-ﾉﾝﾗｲﾄﾌﾟﾛﾃｸﾄ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rcv As String
        Dim SendCom As String = ""
        Dim RecvCom As String = ""
        Dim FuncName As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        For i = 1 To 1
            Select Case i
                'Case 0 : FuncName = "準備：入力ﾎﾟｰﾄの設定" : SendCom = "TS,7,0,0,FFFE398E,2,A0AA," : RecvCom = "TS,7,0,0,0,2,FFFE398E,A0AA,"
                Case 1 : FuncName = "WriteProtect無し確認" : SendCom = "TS,6,0,0,FFFE3982,2," : RecvCom = "TS,6,0,0,0,2,0,FFFE3982,"
            End Select

            Call WriteLog(False, FuncName & "中・・・")
            rcv = TSSendRecv(SendCom, 3)
            If rcv <> RecvCom Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = RecvCom    'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, FuncName & " OK ")
        Next

        ' ---------------------------------------------------------------------------------------
    End Sub

    Public Sub I34(ByVal iname As String)
        '==============================================================================================================
        ' Function : I34:IOﾊﾞｽ検査前準備
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim slot As Integer = 5
        'IOﾊﾞｽﾃｽﾄをｸﾗｽﾓｼﾞｭｰﾙ化大幅修正　        V0.01
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
        Call WriteLog(False, "IOﾊﾞｽﾃｽﾄ前準備中・・・約1秒")

        cBUBU = New cls123_BUBU_IOBUS_TEST
        cBUBU.TESTSlotAddress = slot
        cBUBU.IPAdd = IP_U2

        Try
            Call CountStart(15)
            If cBUBU.TestPreparation() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽﾃｽﾄ前準備失敗"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cBUBU.LOG)
            cBUBU = Nothing
        End Try

        Call WriteLog(False, "IOﾊﾞｽﾃｽﾄ前準備OK")

    End Sub

    Public Sub I35(ByVal iname As String)
        '==============================================================================================================
        ' Function : I35:出荷設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim FuncName As String
        Dim SendCom As String = ""
        Dim RecvCom As String = ""
        Dim rcv As String

        Call WriteLog(False, "")
        Call WriteLog(False, iname & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        FuncName = "ﾛｰﾀﾘｰｽｲｯﾁ0設定指示"
        Call WriteLog(False, FuncName)

        'メッセージ表示
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\RSW1.ppt"
            dfrm.DISP_TITLENAME = iname
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = FuncName & "異常 "
                    mTestSpec = "ﾛｰﾀﾘｰｽｲｯﾁ0設定=0"   'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()
                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            dfrm.Close()
            dfrm = Nothing
        End Try

        For i = 0 To 1
            Select Case i
                Case 0 : FuncName = "ﾛｰﾀﾘｰｽｲｯﾁ0位置読み出し" : SendCom = "TS,6,0,0,FFFE3B02,2," : RecvCom = "TS,6,0,0,0,2,E07F,FFFE3B02,"
                Case 1 : FuncName = "SRAMﾒﾓﾘｸﾘｱ" : SendCom = "TS,52,0,0," : RecvCom = "TS,52,0,0,0,"
            End Select

            Call WriteLog(False, FuncName & "中・・・")
            rcv = TSSendRecv(SendCom, 120)
            If rcv <> RecvCom Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = RecvCom    'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, FuncName & " OK ")
        Next

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

    Public Sub I36(ByVal iname As String)
        '==============================================================================================================
        ' Function : I36:ﾗﾀﾞｰﾌﾗｯｼｭ部出荷設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Dim FuncName As String
        Dim SendCom As String = ""
        Dim RecvCom As String = ""
        Dim rcv As String

        Call WriteLog(False, "")
        Call WriteLog(False, iname & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        FuncName = "ﾗﾀﾞｰﾌﾗｯｼｭﾁｯﾌﾟｲﾚｰｽ" : SendCom = "TS,24,0,0,6,0,0," : RecvCom = "TS,24,0,0,0,6,30000000,FFFFFFFF,FFFFFFFF,"
        If Strings.Left(Model, 6) = "F3SP76" Then Call WriteLog(False, FuncName & "中・・・約40秒")
        If Strings.Left(Model, 6) = "F3SP71" Then Call WriteLog(False, FuncName & "中・・・約20秒")
        If Strings.Left(Model, 6) = "F3SPV9" Then Call WriteLog(False, FuncName & "中・・・約40秒") 'Rev0.12ADD

        Try
            Call CountStart(200)
            rcv = TSSendRecv(SendCom, 90)
            If rcv <> RecvCom Then
                mErrorComment = FuncName & "　異常 " & rcv
                mTestSpec = RecvCom    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
        End Try

        Call WriteLog(False, FuncName & " OK ")

    End Sub

    Public Sub I37(ByVal iname As String)
        '==============================================================================================================
        ' Function : I37:RTC確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim RTCTime As String
        Dim PCTime As String
        Dim RecvDataLE As String = ""

        '-------------------- Ver 0.29 JDI ADD Start --------------------
        Dim dt1 As Date
        Dim dt2 As Date
        Dim SPDate() As Integer = {0, 0, 0, 0, 0, 0}
        Dim result_Time As Long = 0
        '--------------------- Ver 0.29 JDI ADD End ---------------------

        Call WriteLog(False, "")
        Call WriteLog(False, iname & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        '-------------------- Ver 0.29 JDI ADD Start --------------------
        Call WriteLog(True, "XDFail実施")
        cBUS = New cls122_FAM3_IOBUS_Ctrl
        Call cBUS.IOBusTstSetReSetCpuXDFAIL()
        cBUS = Nothing

        'RTC確認 ---
        Call WriteLog(True, "RTC確認")
        'Ver 0.29 RTC確認NG発生。待ち時間5秒->7秒に変更(Debug 1月23日 15:51)
        'Ver 0.29 RTC確認NG発生。待ち時間7秒->10秒に変更(Debug 5月24日)
        'Call WriteLog(False, "読込開始待ち・・・5秒")
        'Call WaitStart(5)
        'Call WriteLog(False, "読込開始待ち・・・7秒")
        'Call WaitStart(7)
        'Call WriteLog(False, "読込開始待ち・・・10秒")       'Ver 0.35 JDI DEL
        'Call WaitStart(10)                                   'Ver 0.35 JDI DEL
        Call WriteLog(False, "読込開始待ち・・・5秒")       'Ver 0.35 JDI ADD
        Call WaitStart(5)                                   'Ver 0.35 JDI ADD
        '--------------------- Ver 0.29 JDI ADD End ---------------------

        '-------------------- Ver 0.29 JDI DEL Start --------------------
        'RTC確認 -------------------------------------
        'Call WriteLog(False, "RTC確認")
        'Call WriteLog(False, "読込開始待ち・・・3秒")
        'Call WaitStart(3)
        '--------------------- Ver 0.29 JDI DEL End ---------------------

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_SP : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 0 : cMcm.TimeOut = 5
        Dim j As Integer = 0
        Do
            Try
                Call cMcm.RTC_READ()        '"RTC読込み"
                RecvDataLE = cMcm.RecvText
            Catch ex As Exception

            End Try

            RTCTime = ""
            '-------------------- Ver 0.29 JDI DEL Start --------------------
            'For i = 17 To 42 Step 2
            '    RTCTime = RTCTime & Chr(CInt("&H" & Mid(RecvDataLE, i, 2)))
            'Next i

            'PCTime = Format(DateTime.Now, "yy/MM/dd") & Format(DateTime.Now, "HH:mm")

            'If RTCTime = PCTime Then Exit Do
            '--------------------- Ver 0.29 JDI DEL End ---------------------

            '-------------------- Ver 0.29 JDI ADD Start --------------------
            dt2 = DateTime.Now
            PCTime = Format(dt2, "yy/MM/dd") & Format(dt2, "HH:mm:ss")

            For i = 17 To 48 Step 2
                RTCTime = RTCTime & Chr(CInt("&H" & Mid(RecvDataLE, i, 2)))
            Next i

            SPDate(0) = CInt(20 & (Mid(RTCTime, 1, 2)))
            SPDate(1) = CInt(Mid(RTCTime, 4, 2))
            SPDate(2) = CInt(Mid(RTCTime, 7, 2))
            SPDate(3) = CInt(Mid(RTCTime, 9, 2))
            SPDate(4) = CInt(Mid(RTCTime, 12, 2))
            SPDate(5) = CInt(Mid(RTCTime, 15, 2))

            dt1 = New DateTime(SPDate(0), SPDate(1), SPDate(2), SPDate(3), SPDate(4), SPDate(5))

            result_Time = DateDiff("S", dt1, dt2)

            If result_Time >= -3 And result_Time <= 3 Then Exit Do
            '--------------------- Ver 0.29 JDI ADD End ---------------------

            Call WriteLog(False, "RTC確認1回目失敗 SP:" & RTCTime & " PC:" & PCTime)

            If j >= 1 Then
                mErrorComment = "RTC確認2回目失敗 SP:" & RTCTime & " PC:" & PCTime
                mTestSpec = "RTC確認"    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            j = j + 1
            Call WriteLog(False, "2回目読込開始待ち・・・10秒")

            Call WaitTimer(10)

            DoEvents()

        Loop

        Call WriteLog(False, "RTC読込 OK : " & RTCTime & " PCTime:" & PCTime)

    End Sub



    Public Sub I40(ByVal iName As String)
        '==============================================================================================================
        ' Function : I40:電源電圧検査始業点検DMM
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
        Call WriteLog(False, "電源電圧検査中・・・")

        Dim ADDevice As String = ""
        Dim RecvDataLE As String = ""
        Dim Result As String
        Dim sValue As String
        Dim dValue As Double
        Dim StdMax As Single
        Dim StdMin As Single
        Dim SigName As String = ""
        Dim mMeas As Single
        Dim i As Integer

        Dim TestStd_VCC_H As Single = 5.15
        Dim TestStd_VCC_L As Single = 4.85
        Dim TestStd_VCCSDI_H As Single = 3.4
        Dim TestStd_VCCSDI_L As Single = 3.0
        Dim TestStd_VCC3_H As Single = 3.46
        Dim TestStd_VCC3_L As Single = 3.14
        Dim TestStd_V1P2_H As Single = 1.26
        Dim TestStd_V1P2_L As Single = 1.14
        Dim TestStd_V1P5_H As Single = 1.57
        Dim TestStd_V1P5_L As Single = 1.43

        Dim cMeas As New cls140_VoltageMeasure

        Try
           
            ''ﾘﾚｰ設定（GND）
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U1
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "509" & ",001,1", RecvDataLE)   'GND側設定
            'If cLnk.ERR_COMMENT <> "" Then
            '    mErrorComment = cLnk.ERR_COMMENT
            '    Throw New ApplicationException(mErrorComment)
            'End If

            For i = 1 To 6
                Select Case i
                    Case 1
                        SigName = "VCC" : ADDevice = "0006" : StdMax = TestStd_VCC_H : StdMin = TestStd_VCC_L
                    Case 2
                        SigName = "VCCSDI" : ADDevice = "0001" : StdMax = TestStd_VCCSDI_H : StdMin = TestStd_VCCSDI_L
                    Case 3
                        SigName = "VCC3P3" : ADDevice = "0002" : StdMax = TestStd_VCC3_H : StdMin = TestStd_VCC3_L
                    Case 4
                        SigName = "V1P2" : ADDevice = "0004" : StdMax = TestStd_V1P2_H : StdMin = TestStd_V1P2_L
                    Case 5
                        SigName = "V1P5" : ADDevice = "0005" : StdMax = TestStd_V1P5_H : StdMin = TestStd_V1P5_L
                    Case 6
                        SigName = "VCCSDI" : ADDevice = "0007" : StdMax = TestStd_VCCSDI_H : StdMin = TestStd_VCCSDI_L
                End Select

                Call WriteLog(False, SigName & "検査中・・・")

                ''ﾘﾚｰ設定ON
                'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "5" & Strings.Right(ADDevice, 2) & ",001,1", RecvDataLE)
                'If cLnk.ERR_COMMENT <> "" Then
                '    mErrorComment = cLnk.ERR_COMMENT
                '    Throw New ApplicationException(mErrorComment)
                'End If
                ''DMM測定
                'Call WaitStart(1)
                'mMeas = DMM_7555_Measure()
                'v0.07　DMM測定
                mMeas = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

                Call WriteLog(False, "DMM測定値 : " & mMeas.ToString & "V")
                Call WriteValueData(SigName & "_DMM", mMeas.ToString)

                Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
                RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
                Result = Strings.Left(RecvDataLE, 2)
                If Result <> "OK" Then
                    mErrorComment = SigName & "測定返信 異常 " & RecvDataLE
                    mTestSpec = "測定返信:OK"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                sValue = Mid(RecvDataLE, 3, 4)
                dValue = CDec("&h" & sValue) / (20000 / 10)

                Call WriteLog(False, "AD測定値 : " & dValue.ToString & "V")
                Call WriteValueData(SigName, dValue.ToString)

                If System.Math.Abs(mMeas - dValue) > 0.05 Then
                    mErrorComment = "ADとDMMの偏差が0.05V以上あります！AD-CH" & Strings.Right(ADDevice, 2)
                    mTestSpec = "偏差が0.05V以下、AD-CH"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                ''ﾘﾚｰ設定OFF
                'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "5" & Strings.Right(ADDevice, 2) & ",001,0", RecvDataLE)
                'If cLnk.ERR_COMMENT <> "" Then
                '    mErrorComment = cLnk.ERR_COMMENT
                '    Throw New ApplicationException(mErrorComment)
                'End If

                Call WaitStart(1)

                Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
                RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
                Result = Strings.Left(RecvDataLE, 2)
                If Result <> "OK" Then
                    mErrorComment = SigName & "測定返信 異常 " & RecvDataLE
                    mTestSpec = "OK"    'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                sValue = Mid(RecvDataLE, 3, 4)
                dValue = CDec("&h" & sValue) / (20000 / 10)

                Call WriteLog(False, "AD測定値 : " & dValue.ToString & "V")

                If dValue > StdMax Or dValue < StdMin Then
                    mErrorComment = SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
                    mTestMeasure = dValue.ToString     'Rev0.20add
                    mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)

            Next
            ' ---------------------------------------------------------------------------------------
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "501" & ",016,0000000000000000", RecvDataLE)
            cLnk = Nothing
            cMeas = Nothing
        End Try

    End Sub

    Public Sub I90(ByVal iName As String)
        '==============================================================================================================
        ' Function : I90:ﾕﾆｯﾄ電源ON
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim slot As Integer = 2

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
        Dim Model As String = ModelRead()
        Dim SendCom As String
        Dim ResStr As String = ""
        Dim wTime As Double = 0

        If InStr(iName, "I91") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If
        '==========================================================================================
        'TESTMODEであればTESTMODE終了後電源OFF
        SendCom = "TS,13,0,0,"
        If Strings.Left(Model, 9) = "F3SP76-7N" Then ResStr = "F3SP76-7N"
        If Strings.Left(Model, 9) = "F3SP71-4N" Then ResStr = "F3SP71-4N"

        If TSSendRecv(SendCom, 1) = "TS,13,0,0,0," & ResStr & "," Then
            'TestMode終了
            Call I06(iName)
        End If

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

    Public Sub IA2(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA2:MACｱﾄﾞﾚｽ重複確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Dim ServerPath As String = ServerPathRead()
        Dim mMACAddress As String = MACAddressRead()
        Dim Model As String = ModelRead()
        Dim mLineSerial As String = ""
        Dim tbName, sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim FolderPath As String = ""

        Select Case Strings.Left(Model, 5)
            Case "F3SP7" : FolderPath = "SP7x"
            Case "F3SP6" : FolderPath = "SP6x"
            Case "F3SPV" : FolderPath = "SP7x"           'Rev0.12ADD
            Case "F3LE0", "F3LE1" : FolderPath = "LE1x"
        End Select

        'MACｱﾄﾞﾚｽ重複確認
        If mMACAddress = "" Then GoTo noMACCheck 'MACｱﾄﾞﾚｽ空欄時は実施しない
        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
            tbName = "tblPastMAC"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & "PastMAC" & " FROM " & tbName & " WHERE PastMAC = '" & mMACAddress & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            Select Case cADO.DB_DATA_COUNT
                Case 0
                    mErrorComment = "MACｱﾄﾞﾚｽが登録されていません！" : Throw New ApplicationException(mErrorComment)
                Case 1
                    Call WriteLog(False, "MACｱﾄﾞﾚｽは1つ登録確認OK")
                Case Else
                    mErrorComment = "MACｱﾄﾞﾚｽが1つ以上存在しています！" : Throw New ApplicationException(mErrorComment)
            End Select
            
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        Exit Sub

noMACCheck:

        Call WriteLog(False, "MACｱﾄﾞﾚｽNoCheck")

    End Sub

    Public Sub I92(ByVal iName As String)
        '==============================================================================================================
        ' Function : I92:検査開始時処理
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rCmd As String = ""

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'DMM Reset =====================================
        Try

            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\START.ppt"
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

    Public Sub I93(ByVal iName As String)
        '==============================================================================================================
        ' Function : I93:検査終了時メッセージ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rCmd As String = ""

        Dim mMACAddress As String = ""
        Dim mLineSerial As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        '最終確認
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        'Dim rField As String = "MacAddress,LineSerial"
        Dim rField As String = "MacAddress,SerialNo"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            mMACAddress = sData(0)
            mLineSerial = sData(1)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

        Try
            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7x\PassWin.ppt"
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

        Call WriteLog(False, "MACｱﾄﾞﾚｽ :" & mMACAddress)
        Call WriteLog(False, "ｼﾘｱﾙ :" & mLineSerial)

        Call WriteLog(False, "*")

        Call WriteLog(False, "●　検査合格時の作業  ======================================")
        '-----rev0.16 del-----
        'Call WriteLog(False, "・上記の各ｱﾄﾞﾚｽNo.とｼｰﾙが一致している事を確認し、")
        'Call WriteLog(False, "  ｼｰﾙをﾁｪｯｸｼｰﾄに貼り付ける。")
        '---------------------
        Call WriteLog(False, "・全てのｹｰﾌﾞﾙ等を外し、検査ﾕﾆｯﾄから製品を外す。")
        Call WriteLog(False, "・主銘板に印字された内容を確認して、製品に貼り付ける。")
        Call WriteLog(False, "・ﾛｰﾀﾘｽｲｯﾁが[0]である事を確認する。")
        Call WriteLog(False, "-以上、お疲れ様でした。(^o^) -")

        '主銘板再印字ボタン有効化 -=========================================================
        If mMACAddress <> "" And mLineSerial <> "" Then
            Insfrm.btnRePrintLabel.Enabled = True
            AddHandler Insfrm.btnRePrintLabel.Click, AddressOf RePrint_Click
        End If

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
    Public Function IA1() As String
        '==============================================================================================================
        ' Function : IA1:Mac Address Port
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim Serial As String = SerialRead()
        Dim ServerPath As String = ServerPathRead()
        Dim Model As String = ModelRead()

        IA1 = ""
        Dim wTime As Double = 0
        'Dim mcAddress As String
        Dim FolderPath As String = ""

        '==========================================================================================

        Select Case Strings.Left(Model, 5)
            'Case "F3SP7" : FolderPath = "SP7x"  'Rev0.12DEL
            Case "F3SP7", "F3SPV" : FolderPath = "SP7x"  'Rev0.12ADD
            Case "F3SP6" : FolderPath = "SP6x"
            Case "F3LE0", "F3LE1" : FolderPath = "LE1x"
        End Select

        'PLASMAｼﾘｱﾙ番号でMACｱﾄﾞﾚｽ登録があるかどうかﾁｪｯｸ
        'Dim dbName As String = mAppPath & "100 System Data\SP7X\MacAddress.mdb"
        Dim dbName As String = ServerPath & "\MACaddress\" & FolderPath & "\MacAddress.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "PastMAC"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        'Dim Model As String = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblPastMAC"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Serial = '" & Serial & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT <> 0 Then
                IA1 = sData(0) : cADO = Nothing
                Call WriteLog(False, "MacAddress : " & sData(0))
                Call WriteData("MacAddress", sData(0))
                Exit Function
            Else                                                    'Ver 0.26   KSJ ADD
                mErrorComment = "MacAddress 発行されません。"       'Ver 0.26   KSJ ADD
                Throw New ApplicationException(mErrorComment)       'Ver 0.26   KSJ ADD
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "MAC Address 異常。" 'Ver 0.26   KSJ ADD
            Call WriteLog(False, mErrorComment)                             'Ver 0.26   KSJ ADD           
            Throw New ApplicationException(mErrorComment)                   'Ver 0.26   KSJ ADD
        Finally
            cADO = Nothing
        End Try
        '-------------------------Ver 0.26 KSJ Del Start --------------------------
        ''MACｱﾄﾞﾚｽ発番
        'Try
        '    'Get MacAddress -----------------------------
        '    cMac = New cls031_MacAddress
        '    cMac.DBFileName = dbName

        '    Call cMac.IssueMacAddress()
        '    If cMac.ERR_MSG <> "" Then
        '        mErrorComment = cMac.ERR_MSG
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    mcAddress = cMac.MacAddress

        '    cMac = Nothing

        '    Call WriteLog(False, "MacAddress : " & mcAddress)
        '    Call WriteData("MacAddress", mcAddress)

        '    'Next Mac Address ==============================================
        '    Call WriteLog(False, "Next MAC ADDRESS ･･･")

        '    'Issue Next MacAddress -------
        '    cMac = New cls031_MacAddress
        '    cMac.DBFileName = dbName
        '    cMac.Serial = Serial

        '    Call cMac.IssueNextMacAddress()
        '    If cMac.ERR_MSG <> "" Then
        '        mErrorComment = cMac.ERR_MSG
        '        Throw New ApplicationException(mErrorComment)
        '    End If

        '    cMac = Nothing

        '    'LOG Disp -----------------------------------
        '    Call WriteLog(False, "Next MAC ADDRESS OK!")

        '    ' ==============================================================

        'Catch ex As Exception
        '    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
        '    mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)

        'Finally
        '    cMac = Nothing

        'End Try

        'IA1 = mcAddress
        '-------------------------Ver  0.26 KSJ Del Del --------------------------
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

    End Sub

    Public Sub LabelPrint(ByVal dbName As String, ByVal rDTkey As String)

        'v0.02新主銘板対応
        Dim cLblPrnt As cls034_LabelPrint

        cLblPrnt = New cls034_LabelPrint
        Try
            cLblPrnt.PrintForm = "SP7x"
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

    Public Sub IP2(ByVal iName As String)
        '==============================================================================================================
        ' Function : IP2:TestPrintLabel
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        '直流マーク確認メッセージ表示
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP7X\直流マーク確認.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "直流マーク確認 異常 "
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

                Call mSecWait(200)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            dfrm.Close()
            dfrm = Nothing
        End Try

        '主銘板の試し打ち
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

        Call LabelTestPrint(dbName, rDTkey)

    End Sub

    Public Sub LabelTestPrint(ByVal dbName As String, ByVal rDTkey As String)

        Dim cLblPrnt As cls034_LabelPrint

        cLblPrnt = New cls034_LabelPrint
        Try
            cLblPrnt.PrintForm = "SP7x"
            Call cLblPrnt.LabelTestPrint(dbName, rDTkey)
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

    Private Function HardRevRead() As String

        '検査型式の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "ProductHardRev"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        HardRevRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            HardRevRead = sData(0)

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

    'UDP通信 TSｺﾏﾝﾄﾞ送受信
    Private Function TSSendRecv(ByVal Command As String, ByVal TimeOut As Integer) As String

        TSSendRecv = ""
        rcvMsg = ""
        Dim remotePort As Integer = 12290
        Dim localPort As Integer = 12290                                    'バインドするローカルポート番号
        'Dim udp As New System.Net.Sockets.UdpClient(remotePort)            'ローカルポート番号localPortにバインドする
        Dim enc As System.Text.Encoding = System.Text.Encoding.UTF8         '文字コードを指定する

        'Dim RemoteIpEndPoint As IPEndPoint = Nothing
        'Dim UdpData As Integer
        Dim sendMsg As String = Command
        Dim TotalWait As Double
        Dim t, t1 As Integer

        usc = New UDPState()
        usc.endPnt = New Net.IPEndPoint(Net.IPAddress.Parse(IP_SP), remotePort)
        usc.udpClt = New Net.Sockets.UdpClient(localPort)

        If Command <> "SP TEST MODE START" Then sendMsg &= vbCrLf

        Dim sendBytes As Byte() = enc.GetBytes(sendMsg)

        Call WriteLog(False, "Send : " & Command)

        usc.udpClt.BeginReceive(AddressOf receivecallback, usc)
        'rcvMsg = ""

        usc.udpClt.Send(sendBytes, sendBytes.Length, IP_SP, remotePort)        'リモートホストを指定してデータを送信する
        t = System.Environment.TickCount

        'データを受信待ち
        Do
            t1 = System.Environment.TickCount
            TotalWait = CInt((t1 - t) / 1000)
            'If rcvMsg <> "" Then Exit Do                    'Ver 0.38 LCH DEL
            If rcvMsg <> "" Then                             'Ver 0.38 LCH ADD
                Call WriteLog(True, "Recv Data : " & rcvMsg) 'Ver 0.38 LCH ADD
                Exit Do                                      'Ver 0.38 LCH ADD
            End If                                           'Ver 0.38 LCH ADD

            If TotalWait >= TimeOut Then
                TSSendRecv = "SP Communication Time Out"
                usc.udpClt.Close()
                usc = Nothing
                Exit Function
            End If

            DoEvents()

        Loop

        If InStr(rcvMsg, vbCrLf) <> 0 Then
            rcvMsg = Strings.Left(rcvMsg, Len(rcvMsg) - Len(vbCrLf))
        End If
        Call WriteLog(False, "Recv : " & rcvMsg)

        'UDP接続を終了
        usc.udpClt.Close()
        usc = Nothing

        TSSendRecv = rcvMsg

    End Function

    'UDP通信 TSｺﾏﾝﾄﾞ送受信
    Private Function TSSendRecv_NonDisp(ByVal Command As String, ByVal TimeOut As Integer) As String

        TSSendRecv_NonDisp = ""
        rcvMsg = ""
        Dim remotePort As Integer = 12290
        Dim localPort As Integer = 12290                                    'バインドするローカルポート番号
        'Dim udp As New System.Net.Sockets.UdpClient(remotePort)            'ローカルポート番号localPortにバインドする
        Dim enc As System.Text.Encoding = System.Text.Encoding.UTF8         '文字コードを指定する

        'Dim RemoteIpEndPoint As IPEndPoint = Nothing
        'Dim UdpData As Integer
        Dim sendMsg As String = Command
        Dim TotalWait As Double
        Dim t, t1 As Integer

        usc = New UDPState()
        usc.endPnt = New Net.IPEndPoint(Net.IPAddress.Parse(IP_SP), remotePort)
        usc.udpClt = New Net.Sockets.UdpClient(localPort)

        If Command <> "SP TEST MODE START" Then sendMsg &= vbCrLf

        Dim sendBytes As Byte() = enc.GetBytes(sendMsg)

        usc.udpClt.BeginReceive(AddressOf receivecallback, usc)
        'rcvMsg = ""

        usc.udpClt.Send(sendBytes, sendBytes.Length, IP_SP, remotePort)        'リモートホストを指定してデータを送信する
        t = System.Environment.TickCount

        'データを受信待ち
        Do
            t1 = System.Environment.TickCount
            TotalWait = CInt((t1 - t) / 1000)
            If rcvMsg <> "" Then Exit Do

            If TotalWait >= TimeOut Then
                TSSendRecv_NonDisp = "SP Communication Time Out"
                usc.udpClt.Close()
                usc = Nothing
                Exit Function
            End If

            DoEvents()

        Loop

        If InStr(rcvMsg, vbCrLf) <> 0 Then
            rcvMsg = Strings.Left(rcvMsg, Len(rcvMsg) - Len(vbCrLf))
        End If

        'UDP接続を終了
        usc.udpClt.Close()
        usc = Nothing

        TSSendRecv_NonDisp = rcvMsg

    End Function


    Public Class UDPState
        Public endPnt As Net.IPEndPoint
        Public udpClt As Net.Sockets.UdpClient
    End Class

    Public Sub receivecallback(ByVal AR As IAsyncResult)

        Try

            Dim LU As Net.Sockets.UdpClient = CType(AR.AsyncState, UDPState).udpClt
            Dim LE As Net.IPEndPoint = CType(AR.AsyncState, UDPState).endPnt

            Dim receiveBytes As Byte() = LU.EndReceive(AR, LE)
            Dim receiveString As String = System.Text.Encoding.UTF8.GetString(receiveBytes)
            Dim receiveAddress As System.Net.IPAddress = LE.Address
            'ここに受信後の処理を記述
            rcvMsg = receiveString

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RePrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '==============================================================================================================
        'Function : 主銘板再印字 
        'Remarks  :
        '==============================================================================================================

        Call IP1("主銘板再印字")

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

    Private Sub WriteValueData(ByVal fName As String, ByVal sData As String)
        '==========================================================================================
        ' Function : Write Value data(F3SP7XData.mdb)
        ' Data     : Field name , data
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim dbName As String = ""
        Dim tbName As String = ""
        Dim sSQL As String = ""
        Dim rsData() As String = {""}
        Dim rdData() As String = {""}
        Dim rfName() As String = {""}
        Dim rdName() As String = {""}

        Dim DataCount As Integer = 0

        dbName = mAppPath & "210 Result Data\Temporary.mdb"
        tbName = "tblResult"
        sSQL = "SELECT StartDateTime,MSCODE,SerialNo FROM " & tbName

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(rfName, rsData)

            DataCount = cADO.DB_DATA_COUNT

        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        If DataCount = 0 Then Exit Sub

        dbName = mAppPath & "210 Result Data\F3SP7XData.mdb"
        tbName = "tblResult"

        'データがあるかどうかチェック
        sSQL = "SELECT * FROM " & tbName & " "
        sSQL &= "WHERE 検査日時='" & rsData(0) & "' AND MSCODE='" & rsData(1) & "' AND SerialNo='" & rsData(2) & "'"
        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(rdName, rdData)
            DataCount = cADO.DB_DATA_COUNT
        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        'Write of data ----------------------------------------------
        If DataCount = 0 Then
            sSQL = "INSERT INTO " & tbName & " (検査日時,MSCODE,SerialNo," & fName & ") "
            sSQL &= "VALUES ('" & rsData(0) & "','" & rsData(1) & "','" & rsData(2) & "','" & sData & "')"
        Else
            sSQL = "UPDATE " & tbName & " SET "
            sSQL &= fName & " = '" & sData & "' "
            sSQL &= "WHERE 検査日時='" & rsData(0) & "' AND MSCODE='" & rsData(1) & "' AND SerialNo='" & rsData(2) & "'"
        End If

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_EXEC_JET()
        Catch ex As Exception
            mErrorComment = cADO.ERR_COMMENT
            If mErrorComment = "" Then mErrorComment = ex.ToString
            Throw New ApplicationException(mErrorComment)
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

    Private Sub mSecWait(ByVal mSec As Integer)
        '==========================================================================================
        ' F.Name   : mWait
        ' Function : mSec Timer
        ' Data     : msec
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait

    End Sub

    Private Function MACAddressRead() As String

        'MACｱﾄﾞﾚｽの読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "MacAddress"            'InspectionModel  MSCODE
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        MACAddressRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            MACAddressRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

    End Function

    'Public Function RS232CSettingRead(ByVal mPort As String, ByRef mSetting As String, ByRef mTerminalChar As String) As Boolean

    '    Dim dbName As String = ""
    '    Dim tbName As String = ""
    '    Dim fName() As String = {""}
    '    Dim sData() As String = {""}
    '    Dim sSQL As String = ""
    '    Dim fCount As Integer = 0
    '    Dim rCount As Integer = 0
    '    Dim SETDATA() As String

    '    If mPort = "" Then
    '        RS232CSettingRead = False
    '        Exit Function
    '    End If

    '    cADO = New cls010_MDB_Operation
    '    dbName = mAppPath & "010 EquipmentData\機器構成一覧.xls"
    '    tbName = "[RS232C$]"
    '    sSQL = "SELECT 設定 , デリミタ FROM " & tbName & " WHERE PORT LIKE '" & mPort & "'"

    '    cADO.DB_PATH = dbName
    '    cADO.TB_NAME = tbName
    '    cADO.DB_SQL = sSQL

    '    cADO.DB_READ_EXCEL_JET(fName, sData)

    '    If cADO.DB_DATA_COUNT = 0 Then
    '        RS232CSettingRead = False
    '        cADO = Nothing
    '        mErrorComment = "機器構成一覧に設定が登録されていません。"
    '        Exit Function
    '    Else
    '        SETDATA = Split(sData(0), "-")
    '        mSetting = SETDATA(0) & "," & SETDATA(1) & "," & SETDATA(2) & "," & SETDATA(3)
    '        Select Case sData(1)
    '            Case "vbcrlf" : mTerminalChar = vbCrLf
    '            Case "vbcr" : mTerminalChar = vbCr
    '            Case "vblf" : mTerminalChar = vbLf
    '        End Select
    '        RS232CSettingRead = True
    '    End If

    '    cADO = Nothing

    'End Function

    'Private Function DMM_7555_Measure() As Single

    '    Dim mSetting As String = ""
    '    Dim mDelim As String = ""

    '    DMM_7555_Measure = Nothing

    '    Try
    '        'DMM設定
    '        cDMM = New cls101_ComRS232C
    '        cDMM.RS_COMM = Insfrm.SerialPort1
    '        cDMM.PORT = rs232cDMMPort
    '        If RS232CSettingRead(rs232cDMMPort, mSetting, mDelim) = False Then
    '            mErrorComment = cDMM.ERR_COMMENT
    '            Throw New ApplicationException(mErrorComment)
    '        End If
    '        cDMM.SETTING = mSetting : cDMM.TERMINAL_CHAR = mDelim
    '        cDMM.PortSetup()
    '        If cDMM.ERR_COMMENT <> "" Then
    '            mErrorComment = cDMM.ERR_COMMENT
    '            Throw New ApplicationException(mErrorComment)
    '        End If

    '        'Header del
    '        cDMM.SendCmd = "COMM:HEAD OFF" : Call cDMM.FuncSendCmd()
    '        If cDMM.ERR_COMMENT <> "" Then
    '            mErrorComment = cDMM.ERR_COMMENT
    '            Throw New ApplicationException(mErrorComment)
    '        End If

    '        'auto range off
    '        cDMM.SendCmd = "SENS:FUNC:VOLT:DC:AUTO OFF" : Call cDMM.FuncSendCmd()
    '        If cDMM.ERR_COMMENT <> "" Then
    '            mErrorComment = cDMM.ERR_COMMENT
    '            Throw New ApplicationException(mErrorComment)
    '        End If

    '        '20V range
    '        cDMM.SendCmd = "SENS:FUNC:VOLT:DC:RANG 2E+1V" : Call cDMM.FuncSendCmd()
    '        If cDMM.ERR_COMMENT <> "" Then
    '            mErrorComment = cDMM.ERR_COMMENT
    '            Throw New ApplicationException(mErrorComment)
    '        End If

    '        'Measure
    '        cDMM.SendCmd = "MEAsure?" : Call cDMM.FuncSendCmd()
    '        If cDMM.ERR_COMMENT <> "" Then
    '            mErrorComment = cDMM.ERR_COMMENT
    '            Throw New ApplicationException(mErrorComment)
    '        End If

    '        'Receve
    '        Call cDMM.FuncRecvCmd()
    '        If cDMM.ERR_COMMENT <> "" Then
    '            mErrorComment = cDMM.ERR_COMMENT
    '            Throw New ApplicationException(mErrorComment)
    '        End If

    '        DMM_7555_Measure = CSng(cDMM.RECEIVE_DATA)

    '    Catch
    '        If mErrorComment = "" Then mErrorComment = "DMM7555測定異常"
    '        Throw New ApplicationException(mErrorComment)

    '    Finally
    '        Call cDMM.PortClose()
    '        cDMM = Nothing

    '    End Try

    'End Function

End Class
