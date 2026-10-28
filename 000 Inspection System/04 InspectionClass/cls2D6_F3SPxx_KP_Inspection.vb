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

Public Class cls2D6_F3SPxx_KP_Inspection
    'Class ================================================
    Dim cADO As cls010_MDB_Operation
    Dim cFop As cls012_File_Operation
    Dim cTxt As cls013_TextFile_Operation
    Dim cLnk As cls080_F3LinkToLE
    Dim cDMM As cls101_ComRS232C
    Dim cISPComm As cls101_ComRS232C
    Dim cWin As cls001_Window_Control
    Dim cMac As cls031_MacAddress
    'Dim cFWir As cls032_FWireAddress
    'Dim cUSB As cls033_USBAddress
    Dim cSer As cls030_LineSerial
    Dim cLbl As cls011_Excel_Operation
    Dim cMcm As cls120_SP_M_Command
    Dim cPCSP As cls060_PCNetSPEED
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl
    Dim cTEST As cls121_SP_TEST_MODE
    Dim cPF As cls125_PF_IOBUS_TEST
    Dim cKP As cls124_KP_IOBUS_TEST
    Dim cSP As cls121_SP_TEST_MODE
    Dim cSP7X As cls2D2_F3SP7x_Inspection

    Dim dfrm As frm500_MessagePopUp

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String
    Dim mTestSpec As String    'Rev0.20Add
    Dim mTestMeasure As String 'Rev0.20Add

    Dim rs232cSPPort As String = "COM5"

    Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3のLE
    Dim IP_U2 As String = "192.168.1.11"        'ﾒｲﾝCPU

    Dim iName As String

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
                'Case "I06" : Call I06(iName)            'I06:TestMode終了
            Case "I07" : Call I07(iName)            'I07:SRAMﾏｰﾁﾝｸﾞﾃｽﾄ
            Case "I08" : Call I08(iName)            'I08:ﾘｱﾙﾀｲﾑｸﾛｯｸﾃｽﾄ
            Case "I09" : Call I09(iName)            'I09:SPUﾃｽﾄ
            Case "I10" : Call I10(iName)            'I10:ｸﾛｯｸﾃｽﾄ
            Case "I11" : Call I11(iName)            'I11:ｼﾘｱﾙI/Oﾃｽﾄ
            Case "I12" : Call I12(iName)            'I12:ﾊﾞｽｴﾗｰﾃｽﾄ
            Case "I13" : Call I13(iName)            'I13:LEDﾃｽﾄ
            Case "I14" : Call I14(iName)            'I14:XSPHLTﾃｽﾄ
            Case "I15" : Call I15(iName)            'I15:XCPHLTﾃｽﾄ
            Case "I16" : Call I16(iName)            'I16:ROMｶｾｯﾄﾃｽﾄ
            Case "I17" : Call I17(iName)            'I17:ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟﾃｽﾄ
            Case "I18" : Call I18(iName)            'I18:ﾗﾀﾞｰﾌﾟﾛｸﾞﾗﾑﾃｽﾄ
            Case "I19" : Call I19(iName)            'I19:電源電圧測定
            Case "I20" : Call I20(iName)            'I20:工場出荷設定
            Case "I21" : Call I21(iName)            'I21:SIO最高速ﾎﾞｰﾚｰﾄ通信確認
            Case "I22" : Call I22(iName)            'I22:SIO最高速ﾎﾞｰﾚｰﾄ通信確認Mｺﾏﾝﾄﾞ

            Case "I50" : Call I50(iName)            'I50:SLOT/XDFAIL TEST
            Case "I51" : Call I51(iName)            'I51:ACS SIGNAL TEST
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
            Case "I92" : Call I92(iName)            'I92:検査開始時処理
            Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save        'Rev 0.26 KSJ ADD
            Case "IP1" : Call IP1(iName)            'IP1:PrintLabel

            Case "IA0" ' Call IA0(iName)            'IA0:ｼﾘｱﾙ番号発番
                'Case "IA3" ' Call IA3(iName)            'IA3:Cable 使用回数確認 I92で実行　Ver 0.21 JDI ADD

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
        Dim recv As String = ""
        Dim writeModel As String = ModelRead()
        Dim mWriteMAC As String = MACAddressRead()
        Dim FuncName As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        '==============================================================================================================

        Try
            'ﾊﾞｽﾓﾆﾀIOﾓｰﾄﾞで再起動
            cBUS = New cls122_FAM3_IOBUS_Ctrl
            '-------------------------------------- Ver 0.24 JDI DEL Start ---------------------------------------------
            'FuncName = "ﾊﾞｽﾓﾆﾀ電源OFF" : Call WriteLog(False, FuncName)
            'Call cBUS.BUS_MONITOR_POWER_OFF()           'ﾊﾞｽﾓﾆﾀ電源OFF
            'Call WriteLog(False, "wait3秒") : Call WaitStart(3)
            'FuncName = "ﾊﾞｽﾓﾆﾀIOﾓｰﾄﾞで電源ON" : Call WriteLog(False, FuncName)
            'Call cBUS.BUS_MONITOR_POWER_ON_IOMODE()     'ﾊﾞｽﾓﾆﾀIOﾓｰﾄﾞで電源ON
            'Call WriteLog(False, "wait13秒") : Call WaitStart(13)
            'FuncName = "ｽﾛｯﾄ6でIOﾓｰﾄﾞで被検査ﾓｼﾞｭｰﾙ電源ON" : Call WriteLog(False, FuncName)
            'Call cBUS.BUSMONITOR_TB04_START()           'ｽﾛｯﾄ6でIOﾓｰﾄﾞで被検査ﾓｼﾞｭｰﾙ電源ON
            'FuncName = "A側SP28ﾘｾｯﾄｽﾀｰﾄ" : Call WriteLog(False, FuncName)
            'Call cBUS.Slot2CpuReset()                   'A側SP28ﾘｾｯﾄｽﾀｰﾄ
            'Call WaitStart(3)
            '--------------------------------------- Ver 0.24 JDI DEL END ----------------------------------------------

            Call cBUS.BUSMONITOR_TB04_END()             'Ver 0.24 JDI ADD
            Call cBUS.BUSMONITOR_TB04_CPU_START(5)      'Ver 0.24 JDI ADD

            FuncName = "画像処理PCからRomPainterでﾌｧｰﾑ書込み" : Call WriteLog(False, FuncName)
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = "192.168.1.30" : cLnk.Port = 1600 : cLnk.TimeOut = 30

            '書き込み形名の指定
            Call WriteLog(False, "RomPainter書き込み形名の指定：" & writeModel)
            Call cLnk.SocketCommunication(writeModel, recv) : Call WriteLog(False, "RomPainter戻り値：" & recv)
            If InStr(recv, "ERR") <> 0 Then mErrorComment = "RomPainterﾌｧｰﾑ形名指定NG" : Throw New ApplicationException(mErrorComment)
            Call WaitStart(3)
            '書き込むMACｱﾄﾞﾚｽの設定
            If mWriteMAC <> "" Then
                Call WriteLog(False, "書き込むMACｱﾄﾞﾚｽ:" & mWriteMAC)
                Call cLnk.SocketCommunication(mWriteMAC, recv) : Call WriteLog(False, "MACｱﾄﾞﾚｽ設定戻り値：" & recv)
                If InStr(recv, "ERR") <> 0 Then mErrorComment = "RomPainterMACｱﾄﾞﾚｽ設定NG" : Throw New ApplicationException(mErrorComment)
            End If
            '消去
            Call WriteLog(False, "RomPainter消去中・・・") : cLnk.TimeOut = 30
            Call cLnk.SocketCommunication("AllERASE", recv) : Call WriteLog(False, "RomPainter戻り値：" & recv)
            If InStr(recv, "ERR") <> 0 Then mErrorComment = "RomPainterﾌｧｰﾑ消去実行NG" : Throw New ApplicationException(mErrorComment)
            Call WaitStart(1)
            '書き込み
            Call WriteLog(False, "RomPainterファーム書き込み中・・・") : cLnk.TimeOut = 120
            Call cLnk.SocketCommunication("WRITE", recv) : Call WriteLog(False, "RomPainter戻り値：" & recv)
            If InStr(recv, "ERR") <> 0 Then mErrorComment = "RomPainterﾌｧｰﾑ書込実行NG" : Throw New ApplicationException(mErrorComment)

            Call WaitStart(3)
            'Call cBUS.BUSMONITOR_TB04_END()                            '検査ﾓｼﾞｭｰﾙ電源OFF Ver 0.24 JDI DEL

        Catch ex As Exception
            If mErrorComment = "" Then
                mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            End If
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            '-------------------------------------- Ver 0.24 JDI DEL Start ---------------------------------------------
            'Call cBUS.BUS_MONITOR_POWER_OFF()
            'Call WaitStart(1)
            'Call cBUS.BUS_MONITOR_POWER_ON()            'CPUﾓｰﾄﾞで電源ON
            'Call WaitStart(13)
            '--------------------------------------- Ver 0.24 JDI DEL END ----------------------------------------------

            cBUS = New cls122_FAM3_IOBUS_Ctrl                          'Ver 0.24 JDI ADD 
            Call cBUS.BUSMONITOR_TB04_END()                            '検査ﾓｼﾞｭｰﾙ電源OFF Ver 0.24 JDI ADD

            cBUS = Nothing
            cLnk = Nothing
        End Try

        Call WriteLog(False, "ﾌｧｰﾑ書込みOK")

    End Sub

    Public Sub I03(ByVal iName As String)
        '==============================================================================================================
        ' Function : I03:ﾓｼﾞｭｰﾙIDﾁｪｯｸ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()
        Dim CardID As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            'ﾓｼﾞｭｰﾙID確認
            Call WriteLog(False, "ﾓｼﾞｭｰﾙID確認")
            cSP.ICMD = 19 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 0
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(5) <> Mid(Model, 3, 4) Then
                'mErrorComment = "ﾓｼﾞｭｰﾙID確認異常。戻り値：" & Mid(Model, 3, 4)     'Rev0.20del
                mErrorComment = "ﾓｼﾞｭｰﾙID確認異常。戻り値：" & RecvDATA(5)     'Rev0.20add
                mTestSpec = "基準値：" & Mid(Model, 3, 4)                             'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, "ﾓｼﾞｭｰﾙID確認OK")
            End If
           
            'CardIDﾃｽﾄ
            Call WriteLog(False, "CardIDﾃｽﾄ")
            cSP.ICMD = 6 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 2
            ReDim SendParam(2)
            SendParam(1) = "60B201" : SendParam(2) = "1" : cSP.SendPARA = SendParam
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            Select Case Mid(Model, 3, 4)
                Case "BP20", "BP30", "MP30"
                    CardID = "40"
                Case "SP22", "SP28", "SP38", "SP53", "SP58", "SP59", "SP35", "SP37", "FP36", "SPV3", "SPV8"
                    CardID = "20"
                Case Else
                    CardID = "10"
            End Select

            If RecvDATA(5) <> CardID Then
                'mErrorComment = "CardID確認異常。戻り値：" & Mid(Model, 3, 4)         'Rev0.20del
                mErrorComment = "CardID確認異常。戻り値：" & RecvDATA(5)                 'Rev0.20add
                mTestSpec = "基準値：" & CardID                              'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, "CardID確認OK")
            End If

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

    End Sub

    Public Sub I04(ByVal iName As String)
        '==============================================================================================================
        ' Function : I04:ROMｻﾑﾁｪｯｸ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()
        'Dim Dummy As String()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim dbName As String
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim ROM_SUM As String = ""

        'If InStr(Model, "/") <> 0 Then Dummy = Split(Model, "/") : Model = Dummy(0)

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

        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            'ROMSUMﾁｪｯｸ
            Call WriteLog(False, "ROMSUMﾁｪｯｸ")
            cSP.ICMD = 4 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 2
            ReDim SendParam(2)
            SendParam(1) = "000000" : SendParam(2) = "0FFFFF" : cSP.SendPARA = SendParam
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(5) <> ROM_SUM Then
                mErrorComment = "ROMSUMﾁｪｯｸ異常。期待値：" & ROM_SUM & "　戻り値：" & RecvDATA(5)
                mTestSpec = "期待値：" & ROM_SUM    'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, "ROMSUMﾁｪｯｸOK")
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
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

        If InStr(iName, "I05") <> 0 Then
            Call WriteLog(False, "")
            Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
            Call WriteLog(False, "*")
        End If

        '==============================================================================================================
        '電源ON
        Call I90("1")                                   '被検査ﾓｼﾞｭｰﾙをﾒｲﾝCPUで電源ON
        If InStr(Model, "F3SP28") <> 0 Or InStr(Model, "F3SP22") <> 0 Then
            Call WaitStart(7)
        Else
            Call WaitStart(3)
        End If

        'SPﾃｽﾄﾓｰﾄﾞ移行
        cSP = New cls121_SP_TEST_MODE
        cSP.PORT = rs232cSPPort
        cSP.RS_COMM = Insfrm.SerialPort1
        Call WriteLog(False, "SPﾃｽﾄﾓｰﾄﾞ移行中。。。しばらくお待ちください。")
        Try
            Call cSP.OldSpTESTMODEStart()
            Call WriteLog(False, cSP.LOG)
        Catch ex As Exception
            mErrorComment = "SPﾃｽﾄﾓｰﾄﾞ移行ｴﾗｰ"
            Call WriteLog(False, cSP.LOG)
            Call WriteLog(False, cSP.ERR_COMMENT)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

        WaitStart(7)
        Call WriteLog(False, "SPﾃｽﾄﾓｰﾄﾞ移行OK")

    End Sub



    Public Sub I07(ByVal iName As String)
        '==============================================================================================================
        ' Function : I07:SRAMﾏｰﾁﾝｸﾞﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            'SRAMﾏｰﾁﾝｸﾞﾃｽﾄ
            Call WriteLog(False, "SRAMﾏｰﾁﾝｸﾞﾃｽﾄ")
            cSP.ICMD = 3 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 3
            ReDim SendParam(3)
            SendParam(1) = "1" : SendParam(2) = "800000" : SendParam(3) = "8FFFFF" : cSP.SendPARA = SendParam
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "3" Or RecvDATA(4) <> "0" Then
                'mErrorComment = "SRAMﾏｰﾁﾝｸﾞﾃｽﾄ異常。"                                     'Rev0.20del
                mErrorComment = "SRAMﾏｰﾁﾝｸﾞﾃｽﾄ異常。 " & RecvDATA(1) & ", " & RecvDATA(4)   'Rev0.20add
                mTestSpec = "3, 0"   'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, "SRAMﾏｰﾁﾝｸﾞﾃｽﾄOK")
            End If

            'SP59のときのみ
            If InStr(Model, "F3SP59-7S") <> 0 Then
                Call WriteLog(False, "F3SP59-7SのみSRAMﾃｽﾄ")
                cSP.ICMD = 3 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 3
                ReDim SendParam(3)
                SendParam(1) = "1" : SendParam(2) = "900000" : SendParam(3) = "9FFFFF" : cSP.SendPARA = SendParam
                Call cSP.TestComSendRecv()
                Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                If RecvDATA(1) <> "3" Or RecvDATA(4) <> "0" Then
                    'mErrorComment = "F3SP59-7SのみSRAMﾃｽﾄ異常。"                                    'Rev0.20del
                    mErrorComment = "F3SP59-7SのみSRAMﾃｽﾄ異常。" & RecvDATA(1) & ", " & RecvDATA(4)   'Rev0.20add
                    mTestSpec = "3, 0"   'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                Else
                    Call WriteLog(False, "F3SP59-7SのみSRAMﾃｽﾄOK")
                End If
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

    End Sub

    Public Sub I08(ByVal iName As String)
        '==============================================================================================================
        ' Function : I08:ﾘｱﾙﾀｲﾑｸﾛｯｸﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            'ﾘｱﾙﾀｲﾑｸﾛｯｸﾃｽﾄ
            Call WriteLog(False, "ﾘｱﾙﾀｲﾑｸﾛｯｸﾃｽﾄ")
            cSP.ICMD = 8 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 0
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "8" Or RecvDATA(4) <> "0" Then
                'mErrorComment = "ﾘｱﾙﾀｲﾑｸﾛｯｸﾃｽﾄ異常。"   'Rev0.20del
                mErrorComment = "ﾘｱﾙﾀｲﾑｸﾛｯｸﾃｽﾄ異常。" & RecvDATA(1) & ", " & RecvDATA(4)   'Rev0.20add
                mTestSpec = "8, 0"   'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, "ﾘｱﾙﾀｲﾑｸﾛｯｸﾃｽﾄOK")
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

    End Sub

    Public Sub I09(ByVal iName As String)
        '==============================================================================================================
        ' Function : I09:SPUﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()
        Dim FuncName As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        For i = 0 To 2
            Try
                cSP = New cls121_SP_TEST_MODE
                cSP.PORT = rs232cSPPort
                cSP.RS_COMM = Insfrm.SerialPort1

                cSP.ICMD = 5 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 4
                ReDim SendParam(4)
                SendParam(1) = "0" : SendParam(2) = "0" : SendParam(3) = "0" : SendParam(4) = "0"

                Select Case i
                    Case 0
                        FuncName = "SPUﾃｽﾄ0（演算器ﾃｽﾄ）"
                    Case 1
                        FuncName = "SPUﾃｽﾄ1（SPU2動作ｸﾛｯｸﾃｽﾄ）"
                        SendParam(1) = "1"
                        If InStr(Model, "F3SP5") <> 0 Or InStr(Model, "F3SPV") <> 0 Then
                            SendParam(2) = "6F00" : SendParam(3) = "7000"
                        Else
                            SendParam(2) = "2B00" : SendParam(3) = "2BC0"
                        End If
                    Case 2
                        FuncName = "SPUﾃｽﾄ2（ﾃﾞﾊﾞｲｽﾊﾞｯｸｱｯﾌﾟ回路におけるﾗｯﾁ信号のﾁｪｯｸ）"
                        If InStr(Model, "F3SP5") <> 0 Or InStr(Model, "F3SPV") <> 0 Then
                            SendParam(1) = "2"
                        Else
                            Exit For
                        End If
                End Select

                Call WriteLog(False, FuncName) : cSP.SendPARA = SendParam
                Call cSP.TestComSendRecv()
                Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                If RecvDATA(1) <> "5" Or RecvDATA(4) <> "0" Then
                    'mErrorComment = FuncName & "異常。"     'Rev0.20del
                    mErrorComment = FuncName & "異常。 " & RecvDATA(1) & ", " & RecvDATA(4)     'Rev0.20add
                    mTestSpec = "5, 0"   'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                Else
                    Call WriteLog(False, FuncName & "OK")
                End If

            Catch ex As Exception
                If mErrorComment = "" Then mErrorComment = FuncName & "異常:" & ex.ToString
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                cSP = Nothing
            End Try

        Next

        Dim DEVREG As String = ""

        For i = 1 To 10
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1
            cSP.IUNIT = 0 : cSP.ISLOT = 0

            Try
                Select Case i
                    Case 1
                        FuncName = "1.ﾃﾞﾊﾞｲｽﾒﾓﾘｱｸｾｽ制御ﾚｼﾞｽﾀをﾘｰﾄﾞし待避する。"
                        If InStr(Model, "F3SP5") <> 0 Or InStr(Model, "F3SPV") <> 0 Then
                            cSP.ICMD = 6 : cSP.IPARA = 2
                            ReDim SendParam(2)
                            SendParam(1) = "F00008" : SendParam(2) = "2" : cSP.SendPARA = SendParam
                            Call WriteLog(False, FuncName)
                            Call cSP.TestComSendRecv()
                            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                            If RecvDATA(1) <> "6" Or RecvDATA(4) <> "0" Then
                                mTestSpec = "6, 0"     'Rev0.20Add
                                Throw New ApplicationException(mErrorComment)
                            Else
                                DEVREG = RecvDATA(5)
                                Call WriteLog(False, FuncName & "OK")
                            End If
                        Else
                            cSP = Nothing : GoTo looploop
                        End If
                    Case 2
                        FuncName = "2.ﾃﾞﾊﾞｲｽﾒﾓﾘｱｸｾｽ制御ﾚｼﾞｽﾀをﾃﾞﾊﾞｲｽﾜｰｸﾘｰﾄﾞ/ﾗｲﾄに設定する。"
                        If InStr(Model, "F3SP5") <> 0 Or InStr(Model, "F3SPV") <> 0 Then
                            cSP.ICMD = 7 : cSP.IPARA = 3
                            ReDim SendParam(3)
                            SendParam(1) = "F00008" : SendParam(2) = "2" : SendParam(3) = "0002" : cSP.SendPARA = SendParam
                            Call WriteLog(False, FuncName)
                            Call cSP.TestComSendRecv()
                            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                        Else
                            cSP = Nothing : GoTo looploop
                        End If
                    Case 3
                        FuncName = "3.ﾃﾞﾊﾞｲｽﾜｰｸﾒﾓﾘのSRAMﾏｰﾁﾝｸﾞﾃｽﾄを行う。"
                        cSP.ICMD = 3 : cSP.IPARA = 3
                        ReDim SendParam(3)
                        SendParam(1) = "1" : SendParam(2) = "C00000"
                        If InStr(Model, "F3SP2") <> 0 Or InStr(Model, "F3SP53") <> 0 Or InStr(Model, "F3SPV3") <> 0 Then SendParam(3) = "C3FFFF"
                        If InStr(Model, "F3SP3") <> 0 Or InStr(Model, "F3SP58") <> 0 Or InStr(Model, "F3SP59") <> 0 _
                           Or InStr(Model, "F3SPV8") <> 0 Then SendParam(3) = "CFFFFF"
                        Call WriteLog(False, FuncName) : cSP.SendPARA = SendParam
                        Call cSP.TestComSendRecv()
                        Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                        Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                        If RecvDATA(1) <> "3" Or RecvDATA(4) <> "0" Then
                            mTestSpec = "3, 0"     'Rev0.20Add
                            Throw New ApplicationException(mErrorComment)
                        Else
                            Call WriteLog(False, FuncName & "OK")
                        End If
                    Case 4
                        FuncName = "4.ﾃﾞﾊﾞｲｽﾒﾓﾘｱｸｾｽ制御ﾚｼﾞｽﾀをﾃﾞﾊﾞｲｽﾊﾞｯｸｱｯﾌﾟﾘｰﾄﾞ/ﾗｲﾄに設定する。"
                        If InStr(Model, "F3SP5") <> 0 Or InStr(Model, "F3SPV") <> 0 Then
                            cSP.ICMD = 7 : cSP.IPARA = 3
                            ReDim SendParam(3)
                            SendParam(1) = "F00008" : SendParam(2) = "2" : SendParam(3) = "0005" : cSP.SendPARA = SendParam
                            Call WriteLog(False, FuncName)
                            Call cSP.TestComSendRecv()
                            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                        Else
                            cSP = Nothing : GoTo looploop
                        End If
                    Case 5
                        FuncName = "5.ﾃﾞﾊﾞｲｽﾊﾞｯｸｱｯﾌﾟﾒﾓﾘのSRAMﾏｰﾁﾝｸﾞﾃｽﾄを行う。"
                        If InStr(Model, "F3SP5") <> 0 Or InStr(Model, "F3SPV") <> 0 Then
                            cSP.ICMD = 3 : cSP.IPARA = 3
                            ReDim SendParam(3)
                            SendParam(1) = "1" : SendParam(2) = "C00000"
                            If InStr(Model, "F3SP53") <> 0 Or InStr(Model, "F3SPV3") <> 0 Then
                                SendParam(3) = "C3FFFF"
                            Else
                                SendParam(3) = "CFFFFF"
                            End If
                            Call WriteLog(False, FuncName) : cSP.SendPARA = SendParam
                            Call cSP.TestComSendRecv()
                            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                            If RecvDATA(1) <> "3" Or RecvDATA(4) <> "0" Then
                                mTestSpec = "3, 0"     'Rev0.20Add
                                Throw New ApplicationException(mErrorComment)
                            Else
                                Call WriteLog(False, FuncName & "OK")
                            End If
                        Else
                            cSP = Nothing : GoTo looploop
                        End If
                    Case 6
                        FuncName = "6.ﾃﾞﾊﾞｲｽﾒﾓﾘｱｸｾｽ制御ﾚｼﾞｽﾀを元に戻す。"
                        If InStr(Model, "F3SP5") <> 0 Or InStr(Model, "F3SPV") <> 0 Then
                            cSP.ICMD = 7 : cSP.IPARA = 3
                            ReDim SendParam(3)
                            SendParam(1) = "F00008" : SendParam(2) = "2" : SendParam(3) = DEVREG : cSP.SendPARA = SendParam
                            Call WriteLog(False, FuncName)
                            Call cSP.TestComSendRecv()
                            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                        Else
                            cSP = Nothing : GoTo looploop
                        End If
                    Case 7
                        FuncName = "7.ﾏｲｸﾛﾌﾟﾛｸﾞﾗﾑﾒﾓﾘのSRAMﾏｰﾁﾝｸﾞﾃｽﾄを行う。"
                        If InStr(Model, "F3SP22") <> 0 Or InStr(Model, "F3SP28-3S") <> 0 Or InStr(Model, "F3SP28-0S") <> 0 Then
                            cSP = Nothing : GoTo looploop
                        Else
                            cSP.ICMD = 3 : cSP.IPARA = 3
                            ReDim SendParam(3)
                            SendParam(1) = "1" : SendParam(2) = "E00000"
                            If InStr(Model, "F3SP53") <> 0 Or InStr(Model, "F3SPV3") <> 0 Then
                                SendParam(3) = "E3FFFF"
                            Else
                                SendParam(3) = "E7FFFF"
                            End If
                            Call WriteLog(False, FuncName) : cSP.SendPARA = SendParam
                            Call cSP.TestComSendRecv()
                            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                            If RecvDATA(1) <> "3" Or RecvDATA(4) <> "0" Then
                                mTestSpec = "3, 0"     'Rev0.20Add
                                Throw New ApplicationException(mErrorComment)
                            Else
                                Call WriteLog(False, FuncName & "OK")
                            End If
                        End If
                    Case 8
                        FuncName = "8.SPU2をRUNさせる。"
                        cSP.ICMD = 7 : cSP.IPARA = 3
                        ReDim SendParam(3)
                        SendParam(1) = "F00002" : SendParam(2) = "2" : SendParam(3) = "0000" : cSP.SendPARA = SendParam
                        Call WriteLog(False, FuncName)
                        Call cSP.TestComSendRecv()
                        Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                        Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                    Case 9
                        FuncName = "9.ｼｰｹﾝｽﾌﾟﾛｸﾞﾗﾑﾒﾓﾘのSRAMﾏｰﾁﾝｸﾞﾃｽﾄを行う。"
                        cSP.ICMD = 3 : cSP.IPARA = 3
                        ReDim SendParam(3)
                        SendParam(1) = "1" : SendParam(2) = "A00000"
                        If InStr(Model, "F3SP2") <> 0 Or InStr(Model, "F3SP53") <> 0 Or InStr(Model, "F3SPV3") <> 0 Then SendParam(3) = "A3FFFF"
                        If InStr(Model, "F3SP58") <> 0 Or InStr(Model, "F3SP59") <> 0 Or InStr(Model, "F3SPV8") <> 0 Then SendParam(3) = "AFFFFF"
                        If InStr(Model, "F3SP3") <> 0 Or InStr(Model, "F3SP28-3N") <> 0 Then SendParam(3) = "A7FFFF"
                        Call WriteLog(False, FuncName) : cSP.SendPARA = SendParam
                        Call cSP.TestComSendRecv()
                        Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                        Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                        If RecvDATA(1) <> "3" Or RecvDATA(4) <> "0" Then
                            mTestSpec = "3, 0"     'Rev0.20Add
                            Throw New ApplicationException(mErrorComment)
                        Else
                            Call WriteLog(False, FuncName & "OK")
                        End If
                    Case 10
                        FuncName = "10.SPU2をｽﾄｯﾌﾟさせる。"
                        cSP.ICMD = 7 : cSP.IPARA = 3
                        ReDim SendParam(3)
                        SendParam(1) = "F00002" : SendParam(2) = "2" : SendParam(3) = "0001" : cSP.SendPARA = SendParam
                        Call WriteLog(False, FuncName)
                        Call cSP.TestComSendRecv()
                        Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                        Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                End Select

            Catch ex As Exception
                If mErrorComment = "" Then mErrorComment = FuncName & "異常:" & ex.ToString
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                cSP = Nothing
            End Try
looploop:
        Next

    End Sub

    Public Sub I10(ByVal iName As String)
        '==============================================================================================================
        ' Function : I10:ｸﾛｯｸﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()
        Dim FuncName As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            FuncName = "ｸﾛｯｸﾃｽﾄ"
            Call WriteLog(False, FuncName)
            cSP.ICMD = 10 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 2
            ReDim SendParam(2)
            SendParam(1) = "E80" : SendParam(2) = "EB0" : cSP.SendPARA = SendParam
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "A" Or RecvDATA(4) <> "0" Then
                'mErrorComment = FuncName & "異常。"   'Rev0.20Del
                mErrorComment = FuncName & "異常。" & RecvDATA(1) & ", " & RecvDATA(4)   'Rev0.20Add
                mTestSpec = "A, 0"     'Rev0.20Add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, FuncName & "OK")
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = FuncName & "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

    End Sub

    Public Sub I11(ByVal iName As String)
        '==============================================================================================================
        ' Function : I11:ｼﾘｱﾙI/Oﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()
        Dim FuncName As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            FuncName = "ｼﾘｱﾙI/Oﾃｽﾄ"
            Call WriteLog(False, FuncName)
            cSP.ICMD = 11 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 0
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "B" Or RecvDATA(4) <> "0" Then
                'mErrorComment = FuncName & "異常。"                                     'Rev0.20Del
                mErrorComment = FuncName & "異常。" & RecvDATA(1) & ", " & RecvDATA(4)   'Rev0.20Add
                mTestSpec = "B, 0"     'Rev0.20Add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, FuncName & "OK")
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = FuncName & "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

    End Sub

    Public Sub I12(ByVal iName As String)
        '==============================================================================================================
        ' Function : I12:ﾊﾞｽｴﾗｰﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()
        Dim FuncName As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            FuncName = "ﾊﾞｽｴﾗｰﾃｽﾄ"
            Call WriteLog(False, FuncName)
            cSP.ICMD = 12 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 0
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "C" Or RecvDATA(4) <> "0" Then
                'mErrorComment = FuncName & "異常。"
                mErrorComment = FuncName & "異常。" & RecvDATA(1) & ", " & RecvDATA(4)   'Rev0.20Add
                mTestSpec = "C, 0"     'Rev0.20Add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, FuncName & "OK")
            End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = FuncName & "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

    End Sub

    Public Sub I13(ByVal iName As String)
        '==============================================================================================================
        ' Function : I13:LEDﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        'Dim RecvDATA As String()
        Dim FuncName As String = ""
        Dim cEI As cls130_EI_GAZO
        Dim fCam As Boolean = False
        'Dim rcv As String
        'Dim Timing As Integer

        Dim NG_Count As Integer           'Ver 0.24 JDI ADD
        Dim Test_Count As Integer         'Ver 0.24 JDI ADD

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        NG_Count = 0                  'Ver 0.24 JDI ADD
        Test_Count = 0                'Ver 0.24 JDI ADD

        '画像処理ﾏｽﾀﾃﾞｰﾀの読込み
        Try
            cEI = New cls130_EI_GAZO
            cEI.IPAdd = "192.168.1.30" : cEI.MasterFile = Mid(Model, 1, 9) & "_NonLED.jpg"
            If cEI.MasterFileRead = False Then
                fCam = False
                mTestSpec = Mid(Model, 1, 9) & "_NonLED.jpg"   'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                fCam = True
            End If
        Catch ex As Exception
            cEI = Nothing
        End Try

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        Select Case fCam
            Case True       'カメラありのとき
                Try
                    'LED-AllOFF
                    Try
                        Call cMcm.IO_WRITE(1, 2, "0000009A", "0000")        'RDY OFF
                    Catch ex As Exception
                    End Try
                    Try
                        Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")        'RUN,ALM OFF
                    Catch ex As Exception
                    End Try
                    Try
                        Call cMcm.IO_WRITE(1, 2, "000000AA", "0000")        'ERR OFF
                    Catch ex As Exception
                    End Try

                    Test_Count += 1                 'Ver 0.24 JDI ADD

                    Call WriteLog(False, "LEDﾌｨﾙﾀﾁｪｯｸ")
                    Call mSecWait(500)
                    Call cEI.GAZO_TEST()
                    If cEI.RESULT <> "0" Then
                        NG_Count += 1                                           'Ver 0.24 JDI ADD           
                        Call WriteLog(False, "ﾌｨﾙﾀﾁｪｯｸ目視検査")                'Ver 0.24 JDI ADD

                        'メッセージ表示
                        dfrm = New frm500_MessagePopUp
                        'dfrm.DISP_TITLENAME = "LEDﾌｨﾙﾀ検査"                                                                  'Rev0.11DEL
                        'dfrm.DISP_MESSAGE = "LEDﾌｨﾙﾀｼﾙｸ印刷は合っていますか？" & vbCrLf & "MODEL=" & Mid(Model, 1, 9)        'Rev0.11DEL
                        dfrm.DISP_TITLENAME = "LED필터 검사"                                                                  'Rev0.11ADD
                        'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까?" & vbCrLf & "MODEL=" & Mid(Model, 1, 9)        'Rev0.11ADD     Ver 0.23 JDI DEL
                        'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까?" & vbCrLf & "MODEL=" & Mid(Model, 3, 7)        'Ver 0.23 ADD /  Ver 0.31 GYS DEL
                        dfrm.DISP_MESSAGE = "필터 형명 확인" & vbCrLf & " = " & Mid(Model, 3, 7)                     'Ver 0.31 GYS ADD
                        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                        Call dfrm.frmShow()

                        Do
                            'wait push button -------------------
                            If dfrm.DIALOG_RESULT = "NG" Then
                                dfrm.Close()
                                dfrm = Nothing
                                mErrorComment = "LEDﾌｨﾙﾀ検査検査 異常 "
                                mTestSpec = "필터검사NG"                'Rev0.20Add
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

                    Call WriteLog(False, "LEDﾌｨﾙﾀﾁｪｯｸOK")

                    cEI.MasterFile = "F3SPxx-xx_LED.jpg"
                    Call cEI.MasterFileRead()

                    For i = 1 To 7

                        Try
                            Select Case i
                                '-----< Rev0.11DEL/ADD >-----
                                'Case 1 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "0000009A", "8080")
                                'Case 2 : FuncName = "RDY,RUN点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "8000")
                                'Case 3 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                'Case 4 : FuncName = "RDY,ALM点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0080")
                                'Case 5 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                'Case 6 : FuncName = "RDY,ERR点灯" : Call cMcm.IO_WRITE(1, 2, "000000AA", "8000")
                                'Case 7 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "000000AA", "0000")
                                Case 1 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "0000009A", "8080")
                                Case 2 : FuncName = "RDY,RUN점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "8000")
                                Case 3 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                Case 4 : FuncName = "RDY,ALM점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0080")
                                Case 5 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                Case 6 : FuncName = "RDY,ERR점등" : Call cMcm.IO_WRITE(1, 2, "000000AA", "8000")
                                Case 7 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "000000AA", "0000")
                                    '----------------------------
                            End Select
                        Catch
                        End Try

                        Test_Count += 1       'Ver 0.24 JDI ADD

                        Call mSecWait(500)
                        Call cEI.GAZO_TEST()

                        Try
                            Select Case FuncName
                                '-----< Rev0.11DEL/ADD >-----
                                'Case "RDY点灯" : If cEI.RESULT <> "0,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,RUN点灯" : If cEI.RESULT <> "0,0,1,1" Then GoTo ErrCheck
                                'Case "RDY,ALM点灯" : If cEI.RESULT <> "0,1,0,1" Then GoTo ErrCheck
                                'Case "RDY,ERR点灯" : If cEI.RESULT <> "0,1,1,0" Then GoTo ErrCheck
                                'Case "全消灯" : If cEI.RESULT <> "1,1,1,1" Then GoTo ErrCheck
                                Case "RDY점등" : If cEI.RESULT <> "0,1,1,1" Then GoTo ErrCheck
                                Case "RDY,RUN점등" : If cEI.RESULT <> "0,0,1,1" Then GoTo ErrCheck
                                Case "RDY,ALM점등" : If cEI.RESULT <> "0,1,0,1" Then GoTo ErrCheck
                                Case "RDY,ERR점등" : If cEI.RESULT <> "0,1,1,0" Then GoTo ErrCheck
                                Case "전소등" : If cEI.RESULT <> "1,1,1,1" Then GoTo ErrCheck
                                    '---------------------------------
                            End Select

                            GoTo NextGo

ErrCheck:
                            NG_Count += 1                                   'Ver 0.24 JDI ADD
                            Call WriteLog(False, "目視検査")                'Ver 0.24 JDI ADD

                            'メッセージ表示
                            dfrm = New frm500_MessagePopUp
                            'dfrm.DISP_TITLENAME = "LED 検査"                         'Rev0.11DEL
                            'dfrm.DISP_MESSAGE = FuncName & "していますか？"          'Rev0.11DEL
                            dfrm.DISP_TITLENAME = "LED 검사"                         'Rev0.11ADD
                            dfrm.DISP_MESSAGE = FuncName & "하고 있습니까?"          'Rev0.11ADD
                            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                            Call dfrm.frmShow()

                            Do
                                'wait push button -------------------
                                If dfrm.DIALOG_RESULT = "NG" Then
                                    dfrm.Close()
                                    dfrm = Nothing
                                    mErrorComment = FuncName & "検査 異常 "
                                    mTestSpec = FuncName            'Rev0.20add
                                    Throw New ApplicationException(mErrorComment)
                                End If

                                If dfrm.DIALOG_RESULT = "OK" Then
                                    Exit Do
                                End If

                                DoEvents()

                            Loop

                            dfrm.Close()
                            dfrm = Nothing

                        Catch ex As Exception
                            cEI = Nothing
                            Call WriteLog(False, mErrorComment)
                            Throw New ApplicationException(mErrorComment)
                        End Try

NextGo:
                        Call WriteLog(False, FuncName & " OK")
                        Call WriteLog(False, cEI.RESULT)

                    Next

                Catch ex As Exception
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                Finally
                    cSP = Nothing
                End Try

            Case False

                'LED検査-------------------------
                Try
                    'メッセージ表示
                    dfrm = New frm500_MessagePopUp
                    dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SPxx\LED.ppt"
                    dfrm.DISP_TITLENAME = "LED 検査"
                    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                    Call dfrm.frmShow()

                    Do
                        For i = 1 To 7
                            Test_Count += 1                                 'Ver 0.24 JDI ADD
                            NG_Count += 1                                   'Ver 0.24 JDI ADD
                            Call WriteLog(False, "目視検査")                'Ver 0.24 JDI ADD

                            Try
                                Select Case i
                                    '-----< Rev0.11DEL/ADD >-----
                                    'Case 1 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "0000009A", "8080")
                                    'Case 2 : FuncName = "RDY,RUN点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "8000")
                                    'Case 3 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                    'Case 4 : FuncName = "RDY,ALM点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0080")
                                    'Case 5 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                    'Case 6 : FuncName = "RDY,ERR点灯" : Call cMcm.IO_WRITE(1, 2, "000000AA", "8000")
                                    'Case 7 : FuncName = "RDY点灯" : Call cMcm.IO_WRITE(1, 2, "000000AA", "0000")
                                    Case 1 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "0000009A", "8080")
                                    Case 2 : FuncName = "RDY,RUN점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "8000")
                                    Case 3 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                    Case 4 : FuncName = "RDY,ALM점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0080")
                                    Case 5 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "000000A8", "0000")
                                    Case 6 : FuncName = "RDY,ERR점등" : Call cMcm.IO_WRITE(1, 2, "000000AA", "8000")
                                    Case 7 : FuncName = "RDY점등" : Call cMcm.IO_WRITE(1, 2, "000000AA", "0000")
                                        '------------------------------------
                                End Select
                            Catch
                            End Try

                            'wait push button -------------------
                            If dfrm.DIALOG_RESULT = "NG" Then
                                mErrorComment = "LED 異常"
                                mTestSpec = FuncName            'Rev0.20add
                                Call WriteLog(False, mErrorComment)
                                Throw New ApplicationException(mErrorComment)
                            End If
                            If dfrm.DIALOG_RESULT = "OK" Then
                                Exit Do
                            End If

                            DoEvents()
                            Call mSecWait(500)
                        Next i
                    Loop

                Catch ex As Exception
                    Throw New ApplicationException(mErrorComment)
                Finally
                    dfrm.Close()
                    dfrm = Nothing
                End Try

                Call WriteLog(False, "LED検査 正常終了")
                ' ---------------------------------------------------------------------------------------

        End Select

        Call WriteLog(False, "検査数：" & Test_Count)                                                    'Ver 0.24 JDI ADD
        Call WriteLog(False, "目視検査数：" & NG_Count)                                                  'Ver 0.24 JDI ADD
        Call WriteLog(False, "自動検査合格率：" & Format((Test_Count - NG_Count) / Test_Count * 100, "0.00") & "%")     'Ver 0.24 JDI ADD

        '全点灯確認
        Try
            'LED全点灯
            Try
                Call cMcm.IO_WRITE(1, 2, "0000009A", "8080")        'RDY ON
                Call cMcm.IO_WRITE(1, 2, "000000A8", "8080")        'RUN,ALM ON
                Call cMcm.IO_WRITE(1, 2, "000000AA", "8000")        'ERR ON
            Catch

            End Try

            Call WriteLog(False, "LED全点灯検査")     'Ver 0.24 JDI ADD
            Call WriteLog(False, "目視検査のみ")      'Ver 0.24 JDI ADD

            'メッセージ表示
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SPXX\ALLLED.ppt"
            dfrm.DISP_TITLENAME = "LED 検査"
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "LED全点灯検査 異常 "
                    mTestSpec = "LED全点灯"            'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

            Loop

        Catch ex As Exception
            mErrorComment = "全点灯LEDｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            cMcm = Nothing
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "LED検査 正常終了")

    End Sub

    Public Sub I14(ByVal iName As String)
        '==============================================================================================================
        ' Function : I14:XSPHLTﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim FuncName As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        'ｽﾛｯﾄ5で電源ON-------------------------------------
        Call I90("ｽﾛｯﾄ5")
        '起動ｳｪｲﾄ7秒
        Call WaitStart(7)
        Try
            cMcm = New cls120_SP_M_Command
            cMcm.IPAdd = IP_U2                      'CPU用LE経由
            cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
            cMcm.SlotNo = 1                         'CPU Slot番号
            cMcm.UnitNo = 0                         'ユニット番号

            FuncName = "SPU2の適当なﾚｼﾞｽﾀ($F00032)に適当な値を書き込む($AA55)"
            Call WriteLog(False, FuncName)
            Call cMcm.IO_WRITE(1, 2, "00000278", "0000")        'IOREGｱﾄﾞﾚｽ0x78に0x0000をﾗｲﾄ
            Call cMcm.IO_WRITE(1, 2, "0000021A", "8080")        'IOREGｱﾄﾞﾚｽ0x1Bに0x80をﾗｲﾄ

            Call cMcm.IO_WRITE(1, 2, "00000210", "00F0")        '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x10）に0x00F0をﾗｲﾄ
            Call cMcm.IO_WRITE(1, 2, "00000212", "0010")        '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x12）に0x0010をﾗｲﾄ

            Call cMcm.IO_WRITE(0, 2, "00000232", "AA55")        'SPU2のｱﾄﾞﾚｽ$F00032に”$AA55”を書き込む
            Call cMcm.IO_READ(0, 2, "00000232")
            Call WriteLog(False, cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 4) <> "AA55" Then
                'mErrorComment = "SPU2ﾚｼﾞｽﾀｱｸｾｽ異常"     'Rev0.20Del
                mErrorComment = "SPU2ﾚｼﾞｽﾀｱｸｾｽ異常, " & Strings.Right(cMcm.RecvText, 4)   'Rev0.20Add
                mTestSpec = "AA55"            'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            FuncName = "ｺﾏﾝﾄﾞﾎﾙﾄを書き込む"
            Call WriteLog(False, FuncName)
            Call cMcm.IO_WRITE(1, 2, "00000210", "0006")        '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x11）に0x06をﾗｲﾄ
            Call cMcm.IO_WRITE(1, 2, "00000212", "2C18")        '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x12）に0x2C18をﾗｲﾄ

            Call cMcm.IO_WRITE(0, 2, "00000201", "0001")        'ｺﾏﾝﾄﾞﾎﾙﾄを設定
            Call cMcm.IO_READ(0, 2, "00000201")                 'ｺﾏﾝﾄﾞﾎﾙﾄを読込み
            Call WriteLog(False, cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "01" Then
                'mErrorComment = "ｺﾏﾝﾄﾞﾎﾙﾄ書き込み異常"       'Rev0.20Del
                mErrorComment = "ｺﾏﾝﾄﾞﾎﾙﾄ書き込み異常, " & Strings.Right(cMcm.RecvText, 2)    'Rev0.20Add
                mTestSpec = "01"            'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            FuncName = "SPU2のﾎﾙﾄｽﾃｰﾀｽ$F00018読込み期待値$0002"
            Call WriteLog(False, FuncName)
            Call cMcm.IO_WRITE(1, 2, "00000210", "00F0")        '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x10）に0x00F0をﾗｲﾄ
            Call cMcm.IO_WRITE(1, 2, "00000212", "0010")        '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x12）に0x0010をﾗｲﾄ

            Call cMcm.IO_READ(0, 2, "00000218")                 'SPU2のｱﾄﾞﾚｽ$F00018に”$0002”が書かれていることを確認
            Call WriteLog(False, cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "02" Then
                'mErrorComment = "SPU2ﾎﾙﾄｽﾃｰﾀｽ確認異常"                                      'Rev0.20del
                mErrorComment = "SPU2ﾎﾙﾄｽﾃｰﾀｽ確認異常, " & Strings.Right(cMcm.RecvText, 2)    'Rev0.20Add
                mTestSpec = "02"            'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            FuncName = "ｺﾏﾝﾄﾞﾘｾｯﾄを書き込む"
            Call WriteLog(False, FuncName)
            Call cMcm.IO_WRITE(1, 2, "00000270", "1000")

            'ｳｪｲﾄ7秒
            Call WaitStart(7)

            FuncName = "CPUが正常に起動しているか確認"
            Call WriteLog(False, FuncName)
            FuncName = "KP上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ$211確認（期待値$80）"
            Call WriteLog(False, FuncName)
            Call cMcm.IO_READ(1, 2, "00000210")
            Call WriteLog(False, cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 2) <> "80" Then
                'mErrorComment = FuncName & "異常"           'Rev0.20del
                mErrorComment = FuncName & "異常, " & Strings.Right(cMcm.RecvText, 2)   'Rev0.20add
                mTestSpec = "80"            'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            FuncName = "KPｽﾚｰﾌﾞｽﾃｰﾀｽ$278確認（期待値$0000）"
            Call WriteLog(False, FuncName)
            Call cMcm.IO_READ(1, 2, "00000278")
            Call WriteLog(False, cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 4) <> "0000" Then
                'mErrorComment = FuncName & "異常"                                     'Rev0.20del
                mErrorComment = FuncName & "異常, " & Strings.Right(cMcm.RecvText, 4)   'Rev0.20add
                mTestSpec = "0000"            'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            FuncName = "SPU2ﾚｼﾞｽﾀ($F00032)に書き込んだ値が初期値（$0000）に戻っていることを確認"
            Call WriteLog(False, FuncName)
            Call cMcm.IO_WRITE(1, 2, "00000210", "00F0")        '上位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x10）に0x00F0をﾗｲﾄ
            Call cMcm.IO_WRITE(1, 2, "00000212", "0010")        '下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽ（0x12）に0x0010をﾗｲﾄ

            Call cMcm.IO_READ(0, 2, "00000232")
            Call WriteLog(False, cMcm.RecvText)
            If Strings.Right(cMcm.RecvText, 4) <> "0000" Then
                'mErrorComment = FuncName & "異常"                                       'Rev0.20del
                mErrorComment = FuncName & "異常, " & Strings.Right(cMcm.RecvText, 4)   'Rev0.20add
                mTestSpec = "0000"            'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

        Catch
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm = Nothing
        End Try

    End Sub

    Public Sub I15(ByVal iName As String)
        '==============================================================================================================
        ' Function : I15:XCPHLTﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        'Dim RecvDATA As String()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            'XCPHLTﾃｽﾄ
            Call WriteLog(False, "XCPHLTﾃｽﾄ")
            cSP.ICMD = 7 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 3
            ReDim SendParam(3)
            SendParam(1) = "60B001" : SendParam(2) = "1" : SendParam(3) = "00000001" : cSP.SendPARA = SendParam

            Call WriteLog(False, "Send : TS,7,0,0,60B001,1,00000001,")
            Call WriteLog(False, "受信異常が表示されても正常です。（約30秒待ち）")
            Call CountStart(90)
            Call cSP.TestComSendRecvTimeout30()

            GoTo ErrTrap

        Catch ex As Exception
            Call WriteLog(False, "受信ﾀｲﾑｱｳﾄ正常。XCPHLTﾃｽﾄOK")
            Call CountStop()
        Finally
            cSP = Nothing
        End Try

        Exit Sub

ErrTrap:
        Call WriteLog(False, "XCPHLTﾃｽﾄ異常")
        Throw New ApplicationException(mErrorComment)

    End Sub


    Public Sub I16(ByVal iName As String)
        '==============================================================================================================
        ' Function : I16:ROMｶｾｯﾄﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RecvDATA As String()
        Dim FuncName As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            FuncName = "ROMｶｾｯﾄﾃｽﾄ"
            Call WriteLog(False, FuncName)
            cSP.ICMD = 9 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 4
            ReDim SendParam(4)
            SendParam(1) = "5" : SendParam(3) = "1" : SendParam(4) = "0"
            Select Case Mid(Model, 3, 7)
                Case "SP38-6S", "SP58-6S", "SP59-7S", "SPV8-6S", "SP22-0S", "SP28-0S", "SP28-3S"            'v0.06でSP22,SP28（-3N以外）を追加
                    SendParam(2) = "8" : Call WriteLog(False, "約60秒待ち")          'AMD 16M
                Case Else
                    SendParam(2) = "7" : Call WriteLog(False, "約30秒待ち")          'AMD 8M
            End Select
            cSP.SendPARA = SendParam
            Call CountStart(90)
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "9" Or RecvDATA(4) <> "0" Then
                'mErrorComment = FuncName & "異常。"                                       'Rev0.20del
                mErrorComment = FuncName & "異常。 " & RecvDATA(1) & ", " & RecvDATA(4)    'Rev0.20add
                mTestSpec = "9, 0"                     'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, FuncName & "OK")
            End If

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
            Call CountStop()
        End Try

    End Sub

    Public Sub I17(ByVal iName As String)
        '==============================================================================================================
        ' Function : I17:ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================

        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()

        Dim FuncName As String = ""
        Dim RecvDataLE As String
        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Dim ADDevice As String = ""
        Dim Result As String
        'Dim sValue As String
        Dim dValue As Single
        Dim StdMax As Single
        Dim StdMin As Single
        Dim SigName As String = ""

        Dim PwrONWait As Integer
        Dim PwrOFFWait As Integer
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Dim cMeas As New cls140_VoltageMeasure

        Dim TestStd_VC3SLP_OFF_H As Single = 3.15           'v0.06 3.1->3.15
        Dim TestStd_VC3SLP_OFF_L As Single = 2.3

        Dim TestStd_VCCSLP_OFF_H As Single = 3.35           'v0.06 3.3->3.35
        Dim TestStd_VCCSLP_OFF_L As Single = 2.3

        Select Case Mid(Model, 3, 4)    'v0.08 VC3SLP_OFF電圧上限
            Case "SP58", "SPV8", "SP59" : TestStd_VC3SLP_OFF_H = 3.4
            Case Else
        End Select

        '検査用ﾒｲﾝCPU(SP7x)からMｺﾏﾝﾄﾞでRTC書き込み。被検査CPUはｱﾄﾞｵﾝでｽﾛｯﾄ2
        'ｽﾛｯﾄ2で電源ON-------------------------------------
        Call I90("ｽﾛｯﾄ2")

        'RTC設定 -------------------------------------
        Call WriteLog(False, "RTC設定")
        Call WriteLog(False, "設定開始待ち・・・5秒")
        Call WaitStart(5)
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2 : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 2 : cMcm.TimeOut = 5
        Dim i As Integer = 0
        Do
            Call cMcm.RTC_SET()     'RTC書き込み(現時刻)
            RecvDataLE = cMcm.RecvText
            If RecvDataLE = "0200021200020000" Then Exit Do '0000021200020000

            If i >= 1 Then
                mErrorComment = iName & "RTC設定　返信　異常 " & RecvDataLE : Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            i = i + 1
            Call WriteLog(False, "RTC設定 ﾘﾄﾗｲ待ち　5秒")
            Call WaitStart(5)
            DoEvents()

        Loop

        '電源OFF
        Call I91(iName)
        PwrOFFWait = 3
        Call WriteLog(False, "電源遮断 " & PwrOFFWait & "秒待ち")
        Call WaitStart(PwrOFFWait)

        '電源OFF電圧
        Try

            For i = 0 To 1
                Select Case i
                    Case 0
                        'VCCSLP-ON -------------------------------------
                        SigName = "VCCSLP_OFF" : ADDevice = "0006"
                        StdMax = TestStd_VCCSLP_OFF_H : StdMin = TestStd_VCCSLP_OFF_L
                    Case 1
                        'VC3SLP-ON -------------------------------------
                        SigName = "VC3SLP_OFF" : ADDevice = "0007"
                        StdMax = TestStd_VC3SLP_OFF_H : StdMin = TestStd_VC3SLP_OFF_L
                End Select

                Call WriteLog(False, "電源OFF時" & SigName & "測定")

                'v0.07　DMM測定
                dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

                'v0.07 DMM測定に変更の為コメント
                'Try
                '    Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
                '    RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
                '    Result = Strings.Left(RecvDataLE, 2)
                '    If Result <> "OK" Then
                '        mErrorComment = "電源OFF時" & SigName & "測定返信 異常 " & RecvDataLE
                '        Call WriteLog(False, mErrorComment)
                '        Throw New ApplicationException(mErrorComment)
                '    End If
                'Catch ex As Exception
                '    Throw New ApplicationException(mErrorComment)
                'Finally
                '    'cLnk = Nothing
                'End Try

                'sValue = Mid(RecvDataLE, 3, 4)
                'dValue = CDec("&h" & sValue) / (20000 / 10)


                If dValue > StdMax Or dValue < StdMin Then
                    mErrorComment = "電源OFF時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString
                    mTestMeasure = dValue.ToString     'Rev0.20add
                    mTestSpec = StdMin.ToString & "～" & StdMax.ToString          'Rev0.20add
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, "電源OFF時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin.ToString & "～" & StdMax.ToString)
                Call WriteValueData(SigName, dValue.ToString)

                ''ﾘﾚｰ設定OFF
                'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "5" & Strings.Right(ADDevice, 2) & ",001,0", RecvDataLE)

                If InStr(Model, "F3SP53") <> 0 Or InStr(Model, "F3SPV3") <> 0 Then Exit For

            Next

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            'Call cLnk.F3PCLinkCommunication("01" & "BWRY" & "501" & ",016,0000000000000000", RecvDataLE)
            'cLnk = Nothing
            'Call cDMM.PortClose()
            'cDMM = Nothing
            cMeas = Nothing
        End Try
        
        'SLPﾗｲﾝを100kΩでGNDに5秒間落とす
        'ﾘﾚｰのｾｯﾄ
        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U1

        SendCom = "01" & "BWRY" & "904" & ",002,11"
        Try
            Call cLnk.F3PCLinkCommunication(SendCom, RecvCom)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cLnk = Nothing
        End Try

        Select Case Mid(Model, 3, 4)    'v0.07で変更ｼｮ-ﾄ時間変更
            Case "SP22", "SP28"
                PwrOFFWait = 5
            Case "SP38", "SP58", "SP59", "SPV8"
                PwrOFFWait = 7
            Case "SP53", "SPV3"
                PwrOFFWait = 1
            Case Else
                PwrOFFWait = 5
        End Select

        Call WriteLog(False, "SLPﾗｲﾝ100kΩ-GND接続 " & PwrOFFWait & "秒待ち")
        Call WaitStart(PwrOFFWait)

        'ﾘﾚｰのﾘｾｯﾄ
        SendCom = "01" & "BWRY" & "904" & ",002,00"
        Try
            Call cLnk.F3PCLinkCommunication(SendCom, RecvCom)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cLnk = Nothing
        End Try

        '電源投入 -------------------------------------
        Call I90(iName)
        PwrONWait = 5
        Call WriteLog(False, "電源投入 " & PwrONWait & "秒待ち")
        Call WaitStart(PwrONWait)

        'RTC確認 -------------------------------------
        Call WriteLog(False, "RTC確認")
        Dim RTCTime, pctime As String

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2 : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 2 : cMcm.TimeOut = 5
        Dim j As Integer = 0
        Do
            Try
                Call cMcm.RTC_READ()        '"RTC読込み"
                RecvDataLE = cMcm.RecvText
            Catch ex As Exception

            End Try

            RTCTime = ""
            For i = 17 To 42 Step 2
                RTCTime = RTCTime & Chr(CInt("&H" & Mid(RecvDataLE, i, 2)))
            Next i
            pctime = Format(DateTime.Now, "yy/MM/dd") & Format(DateTime.Now, "HH:mm")
            If RTCTime = pctime Then Exit Do
            Call WriteLog(False, "RTC確認1回目失敗 SP:" & RTCTime & " PC:" & pctime)

            If j >= 1 Then
                mErrorComment = "RTC確認2回目失敗 SP:" & RTCTime & " PC:" & pctime
                mTestSpec = " PC:" & pctime                        'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            j = j + 1
            Call WriteLog(False, "2回目読込開始待ち・・・2秒")

            Call WaitStart(2)
            DoEvents()

        Loop

        'ResultData.RTC_Read = RecvDataLE
        Call WriteLog(False, "RTC読込 OK : " & RTCTime & " PCTime:" & pctime)

        'ﾊﾞｯﾃﾘｴﾗｰﾁｪｯｸ特殊ﾘﾚｰM194確認
        Call WriteLog(False, "")
        Call WriteLog(False, "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194確認")

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U2
            Call cLnk.F3PCLinkCommunication("02" & "BRD" & "M194" & ",001", RecvDataLE)

            Result = Strings.Mid(RecvDataLE, 3, 2)
            If Result <> "OK" Then
                mErrorComment = "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194　異常 " & RecvDataLE
                mTestSpec = "OK"      'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
            If Mid(RecvDataLE, 5, 1) <> "0" Then
                mErrorComment = "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194異常 " & Mid(RecvDataLE, 3, 4) & " (正常は0)"
                mTestSpec = "正常値：0"      'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "ﾊﾞｯﾃﾘ異常ﾘﾚｰM194=0確認 OK")

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

    End Sub

    Public Sub I18(ByVal iName As String)
        '==============================================================================================================
        ' Function : I18:ﾗﾀﾞｰﾌﾟﾛｸﾞﾗﾑﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim Model As String = ModelRead()
        Dim RecvDataLE As String = ""
        Dim Result As String = ""
        Dim FuncName As String

        '電源OFF
        Call I91(iName)

        'ROMﾊﾟｯｸｾｯﾄﾒｯｾｰｼﾞ
        dfrm = New frm500_MessagePopUp
        'dfrm.DISP_TITLENAME = "ﾗﾀﾞｰ検査"                        'Rev0.11DEL
        'dfrm.DISP_MESSAGE = Mid(Model, 1, 9) & "用のROMﾊﾟｯｸを接続してください。" & vbCrLf & "準備ができたらOKを押してください。"       'Rev0.11DEL
        dfrm.DISP_TITLENAME = "Ladder검사"                        'Rev0.11ADD
        dfrm.DISP_MESSAGE = Mid(Model, 1, 9) & "용 ROM PACK을 접속해 주세요." & vbCrLf & "접속한후에 OK를 눌러주세요."       'Rev0.11ADD
        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        Call dfrm.frmShow()

        Do
            'wait push button -------------------
            If dfrm.DIALOG_RESULT = "NG" Then
                dfrm.Close()
                dfrm = Nothing
                mErrorComment = "ROMﾊﾟｯｸｾｯﾄ 異常 "
                Throw New ApplicationException(mErrorComment)
            End If

            If dfrm.DIALOG_RESULT = "OK" Then
                Exit Do
            End If

            DoEvents()

        Loop
        dfrm.Close()
        dfrm = Nothing

        'ｽﾛｯﾄ2で電源ON
        Call I90("ｽﾛｯﾄ2")

        Select Case Mid(Model, 1, 6)
            Case "F3SP59"
                FuncName = "ｳｪｲﾄ5秒" : Call WriteLog(False, FuncName) : Call WaitStart(5)
            Case Else
                FuncName = "ｳｪｲﾄ3秒" : Call WriteLog(False, FuncName) : Call WaitStart(3)
        End Select

        FuncName = "PCﾘﾝｸｺﾏﾝﾄﾞで内部ﾘﾚｰI001確認（期待値$F00F）" : Call WriteLog(False, FuncName)
        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U2
            Call cLnk.F3PCLinkCommunication("02" & "WRD" & "I001" & ",01", RecvDataLE)

            Call WriteLog(False, RecvDataLE)
            Result = Strings.Mid(RecvDataLE, 3, 2)
            If Result <> "OK" Then
                mErrorComment = "内部ﾘﾚｰI001返信異常 " & RecvDataLE
                mTestSpec = "基準値：OK"     'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
            If Mid(RecvDataLE, 5, 4) <> "F00F" Then
                mErrorComment = "内部ﾘﾚｰI001返信異常 " & Mid(RecvDataLE, 5, 4) & " (正常はF00F)"
                mTestSpec = "正常:F00F"         'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "ﾗﾀﾞｰ検査 OK")

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        '運転停止 ----------------------------------------------
        Call WriteLog(False, "ﾗﾀﾞｰ運転 停止")
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2
        cMcm.UnitNo = 0 : cMcm.SlotNo = 2
        cMcm.RemPort = 12290
        Call cMcm.SEQ_STOP()
        If cMcm.RecvText <> "0200200200020000" Then     '0000200200020000
            mErrorComment = "運転停止指令　異常 " & cMcm.RecvText
            mTestSpec = "基準値：0200200200020000"         'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "ﾗﾀﾞｰ運転 停止 OK")

        Call mSecWait(500)

        'プログラム消去 ----------------------------------------------
        Call WriteLog(False, "ﾗﾀﾞｰ消去")
        Try
            Call cMcm.MEM_CLEAR()
        Catch ex As Exception
        Finally
            cMcm = Nothing
        End Try

        Call WriteLog(False, "ﾗﾀﾞｰ消去 OK")
        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I19(ByVal iName As String)
        '==============================================================================================================
        ' Function : I19:電源電圧測定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Dim ADDevice As String = ""
        'Dim Result As String
        'Dim sValue As String
        Dim dValue As Single
        Dim StdMax As Single
        Dim StdMin As Single
        Dim SigName As String = ""
        'Dim mMeas As Single
        Dim cMeas As New cls140_VoltageMeasure

        Dim i As Integer
        Dim SendCom As String = ""
        Dim RecvCom As String = ""
        Dim RecvDataLE As String = ""

        Dim TestStd_VC3SLP_ON_H As Single = 3.4
        Dim TestStd_VC3SLP_ON_L As Single = 3.2

        Dim TestStd_VCCSLP_ON_H As Single = 5.15        '?暫定
        Dim TestStd_VCCSLP_ON_L As Single = 4.85        '?暫定

        Dim TestStd_VCC_H As Single = 5.15
        Dim TestStd_VCC_L As Single = 4.85

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Dim mVCCSLP As Double

        For i = 0 To 2
            Select Case i
                Case 0
                    'VCCSLP-ON -------------------------------------
                    SigName = "VCCSLP_ON" : ADDevice = "0006"
                    StdMax = TestStd_VCCSLP_ON_H : StdMin = TestStd_VCCSLP_ON_L
                Case 1
                    'VC3SLP-ON -------------------------------------
                    SigName = "VC3SLP_ON" : ADDevice = "0007"
                    StdMax = TestStd_VC3SLP_ON_H : StdMin = TestStd_VC3SLP_ON_L
                    'Case 2
                    '    'VCC -------------------------------------
                    '    SigName = "VCC" : ADDevice = "0008"
                    '    StdMax = TestStd_VCC_H : StdMin = TestStd_VCC_L
                Case 2
                    'PCPort5V -------------------------------------
                    SigName = "PCPort5V" : ADDevice = "0005"
                    StdMax = TestStd_VCC_H : StdMin = TestStd_VCC_L

            End Select

            Call WriteLog(False, "電源ON時" & SigName & "測定")

            'v0.07　DMM測定
            dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))
            'v0.07 DMM測定に変更の為コメント
            'Try
            '    cLnk = New cls080_F3LinkToLE
            '    cLnk.IPAdd = IP_U1
            '    Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
            '    RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
            '    Result = Strings.Left(RecvDataLE, 2)
            '    If Result <> "OK" Then
            '        mErrorComment = "電源ON時" & SigName & "測定返信 異常 " & RecvDataLE
            '        Call WriteLog(False, mErrorComment)
            '        Throw New ApplicationException(mErrorComment)
            '    End If
            'Catch ex As Exception
            '    Throw New ApplicationException(mErrorComment)
            'Finally
            '    cLnk = Nothing
            'End Try

            'sValue = Mid(RecvDataLE, 3, 4)
            'dValue = CDec("&h" & sValue) / (20000 / 10)

            If dValue > StdMax Or dValue < StdMin Then
                mErrorComment = "電源ON時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
                mTestMeasure = dValue.ToString     'Rev0.20add
                mTestSpec = StdMin.ToString & "～" & StdMax.ToString     'Rev0.20add
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)
            Call WriteValueData(SigName, dValue.ToString)

            If i = 0 Then mVCCSLP = dValue
            If i = 2 Then
                If Math.Abs(dValue - mVCCSLP) > 0.2 Then
                    mErrorComment = "VCCとVCCCSLPとの電位差が0.2V以上で異常"
                    mTestMeasure = Math.Abs(dValue - mVCCSLP).ToString      'Rev0.20add
                    mTestSpec = "0.2"    'Rev0.20add
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If
            End If
        Next

    End Sub

    Public Sub I20(ByVal iName As String)
        '==============================================================================================================
        ' Function : I20:工場出荷設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'ｽﾛｯﾄ5で電源ON-------------------------------------
        Call I90("ｽﾛｯﾄ5")

        Call WaitStart(5)

        '電源OFF
        Call I91(iName)

    End Sub

    Public Sub I21(ByVal iName As String)
        '==============================================================================================================
        ' Function : I21:SIO最高速ﾎﾞｰﾚｰﾄ通信確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim recvData As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'ｽﾛｯﾄ2で電源ONしてあること
        Call WriteLog(False, "SIOﾎﾞｰﾚｰﾄ変更115200bps")
        cKP = New cls124_KP_IOBUS_TEST
        cKP.TESTSlotAddress = 2
        cKP.IPAdd = IP_U2

        Try
            Call CountStart(15)
            If cKP.TestPreparation() = False Then
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽﾃｽﾄ前準備失敗"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            cKP = Nothing
        End Try

        'ｼﾘｱﾙ通信
        'SP COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "115200,E,8,1" : cISPComm.TERMINAL_CHAR = Chr(3)
        cISPComm.PortSetup()

        '内部ﾘﾚｰI00001読み出し
        Try
            Call WriteLog(False, "Send : " & "01021WRDI00001,1")
            cISPComm.SendCmd = Chr(2) & "01021WRDI00001,1" & Chr(3)
            Call cISPComm.FuncSendRecvCmd()
            recvData = cISPComm.RECEIVE_DATA
            If InStr(recvData, Chr(2)) = 0 Or InStr(recvData, Chr(2)) = 0 Then
                mErrorComment = "STXorETX受信異常"
                Throw New ApplicationException(mErrorComment)
            End If
            recvData = Replace(recvData, Chr(2), "")
            recvData = Replace(recvData, Chr(3), "")
            Call WriteLog(False, "Recv : " & recvData)
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "SIO最高速ﾎﾞｰﾚｰﾄ通信異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "SIO最高速ﾎﾞｰﾚｰﾄ通信確認正常")

    End Sub

    Public Sub I22(ByVal iName As String)
        '==============================================================================================================
        ' Function : I22:SIO最高速ﾎﾞｰﾚｰﾄ通信確認Mｺﾏﾝﾄﾞ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        'Dim recvData As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        '
        'Mｺﾏﾝﾄﾞ(0401)でSIO最高速ﾎﾞｰﾚｰﾄ通信確認
        cSP = New cls121_SP_TEST_MODE
        cSP.PORT = rs232cSPPort
        cSP.RS_COMM = Insfrm.SerialPort1

        Try
            Call WriteLog(False, "Mｺﾏﾝﾄﾞ(0401)SIO通信速度変更")
            If cSP.ComMode_Change("KP") = False Then
                mErrorComment = "Mｺﾏﾝﾄﾞ(0401)SIO最高速通信速度変更異常"
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, cSP.LOG)

            Call WaitStart(1)   '待ち

            Call WriteLog(False, "Mｺﾏﾝﾄﾞ(0402)SIO最高速通信速度確認")
            Call WriteLog(False, "KP:115200bpsﾊﾟﾘﾃｨなし通信確認")
            If cSP.ComMode_Set("KP") = False Then
                mErrorComment = "Mｺﾏﾝﾄﾞ(0401)SIO最高速通信速度変更異常"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cSP.LOG)

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "Mｺﾏﾝﾄﾞ(0401)でSIO最高速ﾎﾞｰﾚｰﾄ通信確認異常"
            'Call WriteLog(False, cSP.LOG)
            Call WriteLog(False, cSP.ERR_COMMENT)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

        Call WriteLog(False, "SIO最高速ﾎﾞｰﾚｰﾄ通信確認正常")

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
            If cKP.SlotAddressTestForOldSP() = False Then
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, cKP.LOG)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "XACSﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "XBCLRﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "XBRﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "XIRQﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "XLFAILﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "XPFAILﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "XSFAILﾃｽﾄ異常:" & cKP.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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
                Call WriteLog(False, cKP.LOG)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, cKP.LOG)
        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽﾃｽﾄ前準備失敗"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
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

        Call WaitStart(5)
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
        ' Function : I92:検査開始時処理
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rCmd As String = ""
        Dim mDispFileName As String
        Dim MODEL As String = ModelRead()

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        '表示ファイル名選択 =====================================
        Select Case Mid(MODEL, 3, 7)
            Case "SP38-6S", "SP58-6S", "SP59-7S", "SPV8-6S", "SP22-0S", "SP28-0S", "SP28-3S"
                mDispFileName = "SRAM3start.ppt"          'AMD 16M
            Case Else
                mDispFileName = "SRAM2start.ppt"          'AMD 8M
        End Select
        
        Try

            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SPXX\" & mDispFileName
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

            IA3(iName) 'Ver 0.21 JDI ADD Cable交換回数確認。

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
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SPxx\Fin.ppt"
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

        Insfrm.btnRePrintLabel.Enabled = True
        AddHandler Insfrm.btnRePrintLabel.Click, AddressOf RePrint_Click

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
    Public Sub IA0(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA0:LineSerial
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Dim LSerial As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
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

    End Sub

    '--------------------------------------------------- Ver 0.21 JDI ADD Start -------------------------------------------------
    Public Sub IA3(ByVal iName As String)
        '==============================================================================================================
        ' Function : IA3:Cable 使用回数確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim cSPPF As cls2D4_F3SPxx_PF_Inspection
        cSPPF = New cls2D4_F3SPxx_PF_Inspection

        Try
            Call cSPPF.IA3(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub
    '---------------------------------------------------- Ver 0.21 JDI ADD End --------------------------------------------------

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
            cLblPrnt.PrintForm = "STD"
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
        ' Function : Write Value data(F3RP6XData.mdb)
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

        dbName = mAppPath & "210 Result Data\F3SPxxData.mdb"
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

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
    End Sub

End Class
