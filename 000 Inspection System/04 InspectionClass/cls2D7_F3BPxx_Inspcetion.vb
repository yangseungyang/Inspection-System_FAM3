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

Public Class cls2D7_F3BPxx_Inspcetion

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
    Dim cPF As cls125_PF_IOBUS_TEST
    Dim cSPKP As cls2D6_F3SPxx_KP_Inspection
    Dim cSPPF As cls2D4_F3SPxx_PF_Inspection

    Dim cSP As cls121_SP_TEST_MODE

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
    Dim IP_U2 As String = "192.168.1.11"        'ﾒｲﾝCPUのLE
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
                'Case "I09" : Call I09(iName)            'I09:SPUﾃｽﾄ
            Case "I10" : Call I10(iName)            'I10:ｸﾛｯｸﾃｽﾄ
            Case "I11" : Call I11(iName)            'I11:ｼﾘｱﾙI/Oﾃｽﾄ
            Case "I12" : Call I12(iName)            'I12:ﾊﾞｽｴﾗｰﾃｽﾄ
            Case "I13" : Call I13(iName)            'I13:LEDﾃｽﾄ
                'Case "I14" : Call I14(iName)            'I14:XSPHLTﾃｽﾄ
            Case "I15" : Call I15(iName)            'I15:XCPHLTﾃｽﾄ
            Case "I16" : Call I16(iName)            'I16:ROMｶｾｯﾄﾃｽﾄ
            Case "I17" : Call I17(iName)            'I17:ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟﾃｽﾄ
            Case "I18" : Call I18(iName)            'I18:RTC時刻ｶｳﾝﾄﾃｽﾄ
            Case "I19" : Call I19(iName)            'I19:電源電圧測定
            Case "I20" : Call I20(iName)            'I20:工場出荷設定
            Case "I21" : Call I21(iName)            'I21:SIO最高速ﾎﾞｰﾚｰﾄ通信確認
            Case "I22" : Call I22(iName)            'I22:IOﾊﾞｽﾃｽﾄ-XLFAIL(TEST MODE)
            Case "I23" : Call I23(iName)            'I23:SIO最高速ﾎﾞｰﾚｰﾄ通信確認

            Case "I50" : Call I50(iName)            'I50:SLOT/XDFAIL TEST
            Case "I51" : Call I51(iName)            'I51:ACS SIGNAL TEST
            Case "I52" : Call I52(iName)            'I52:BCLR TEST
            Case "I53" : Call I53(iName)            'I53:BR TEST
            Case "I54" : Call I54(iName)            'I54:I/O BUS WRITE/READ TEST
            Case "I55" : Call I55(iName)            'I55:I/O BUS INTERRUPT TEST
            Case "I56" : Call I56(iName)            'I56:XLFAIL TEST
            Case "I57" : Call I57(iName)            'I57:XPFAIL TEST
            Case "I58" : Call I58(iName)            'I58:XSFAIL TEST
            Case "I59" : Call I59(iName)            'I59:IOﾊﾞｽ検査前準備

            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF
            Case "I92" : Call I92(iName)            'I92:検査開始時処理
            Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save        'Rev 0.26 KSJ ADD
            Case "IA0"
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
            mTestSpec = cSPKP.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

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

        If Mid(Model, 1, 6) = "F3BP30" Then
            Try
                cSP = New cls121_SP_TEST_MODE
                cSP.PORT = rs232cSPPort
                cSP.RS_COMM = Insfrm.SerialPort1

                cSP.ICMD = 7 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 3
                ReDim SendParam(3)
                SendParam(1) = "901078" : SendParam(2) = "4" : SendParam(3) = "100000"
                cSP.SendPARA = SendParam
                Call cSP.TestComSendRecv()
                Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                If RecvDATA(1) <> "7" Or RecvDATA(4) <> "0" Then
                    'mErrorComment = "F3BP30-0NﾓｼﾞｭｰﾙIDﾁｪｯｸ準備異常。"        'Rev0.20del
                    mErrorComment = "F3BP30-0NﾓｼﾞｭｰﾙIDﾁｪｯｸ準備異常。 " & RecvDATA(1) & ", " & RecvDATA(4)        'Rev0.20add
                    mTestSpec = "7, 0"                                       'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If

            Catch ex As Exception
                If mErrorComment = "" Then mErrorComment = "異常:" & ex.ToString
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                cSP = Nothing
            End Try

        End If

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I03(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            mTestSpec = cSPKP.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
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
        Dim REV As String = FirmRevRead()

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I04(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

        '+ ROM REV CHECK TEST
        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            Call WriteLog(False, "ROM REV CHECK TEST")

            cSP.ICMD = 6 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 2
            ReDim SendParam(2)

            Select Case Mid(Model, 3, 4)
                Case "BP20"
                    SendParam(1) = "2E612" : SendParam(2) = "2"
                Case "BP30"
                    SendParam(1) = "2E612" : SendParam(2) = "2"
            End Select

            cSP.SendPARA = SendParam
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "6" Or RecvDATA(4) <> "0" Then
                'mErrorComment = "ROM REV CHECK TEST異常。"                                          'Rev0.20del
                mErrorComment = "ROM REV CHECK TEST異常。 " & RecvDATA(1) & ", " & RecvDATA(4)      'Rev0.20add
                mTestSpec = "6, 0"                         'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                If CStr(Chr(CInt("&H" & Mid(RecvDATA(5), 1, 2)))) & CStr(Chr(CInt("&H" & Mid(RecvDATA(5), 3, 2)))) <> REV Then
                    'mErrorComment = "ROM REV CHECK TEST異常。"                                          'Rev0.20del
                    mErrorComment = "ROM REV CHECK TEST異常。 " & CStr(Chr(CInt("&H" & Mid(RecvDATA(5), 1, 2)))) & CStr(Chr(CInt("&H" & Mid(RecvDATA(5), 3, 2))))      'Rev0.20add
                    mTestSpec = REV                         'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                End If
                Call WriteLog(False, "ROM REV CHECK TEST OK")
                End If

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

    End Sub

    Public Sub I05(ByVal iName As String)
        '==============================================================================================================
        ' Function : I05:TestMode移行
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I05(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            mTestSpec = cSPKP.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

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
        Dim FuncName As String = ""

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
            SendParam(1) = "1"

            FuncName = "CPU WORK AREA"
            Select Case Mid(Model, 3, 4)
                Case "BP20"
                    SendParam(2) = "8C0000" : SendParam(3) = "8FFFFF"
                Case "BP30"
                    SendParam(2) = "900000" : SendParam(3) = "9FFFFF"
            End Select
           
            Call WriteLog(False, FuncName)
            cSP.SendPARA = SendParam
            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            If RecvDATA(1) <> "3" Or RecvDATA(4) <> "0" Then
                'mErrorComment = FuncName & ",SRAMﾏｰﾁﾝｸﾞﾃｽﾄ異常。"      'Rev0.20del
                mErrorComment = FuncName & ",SRAMﾏｰﾁﾝｸﾞﾃｽﾄ異常。" & RecvDATA(1) & ", " & RecvDATA(4)      'Rev0.20add
                mTestSpec = "3, 0"                         'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            Else
                Call WriteLog(False, FuncName & "OK")
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

        cSPKP = New cls2D6_F3SPxx_KP_Inspection
        Try
            Call cSPKP.I08(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            mTestSpec = cSPKP.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I10(ByVal iName As String)
        '==============================================================================================================
        ' Function : I10:ｸﾛｯｸﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I10(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I11(ByVal iName As String)
        '==============================================================================================================
        ' Function : I11:ｼﾘｱﾙI/Oﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        cSPKP = New cls2D6_F3SPxx_KP_Inspection

        Try
            Call cSPKP.I11(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            mTestSpec = cSPKP.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I12(ByVal iName As String)
        '==============================================================================================================
        ' Function : I12:ﾊﾞｽｴﾗｰﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        cSPKP = New cls2D6_F3SPxx_KP_Inspection

        Try
            Call cSPKP.I12(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            mTestSpec = cSPKP.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I13(ByVal iName As String)
        '==============================================================================================================
        ' Function : I13:LEDﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I13(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub


    Public Sub I15(ByVal iName As String)
        '==============================================================================================================
        ' Function : I15:XCPHLTﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        cSPKP = New cls2D6_F3SPxx_KP_Inspection

        Try
            Call cSPKP.I15(iName)
        Catch ex As Exception
            mErrorComment = cSPKP.ERR_COMMENT
            mTestSpec = cSPKP.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPKP = Nothing
        End Try

    End Sub

    Public Sub I16(ByVal iName As String)
        '==============================================================================================================
        ' Function : I16:ROMｶｾｯﾄﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I16(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
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

        Dim ADDevice As String = ""
        'Dim Result As String
        Dim SigName As String = ""
        Dim PwrONWait As Integer
        Dim PwrOFFWait As Integer
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

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
                mErrorComment = iName & "RTC設定　返信　異常 " & RecvDataLE
                mTestSpec = "0200021200020000"                       'Rev0.20add
                Throw New ApplicationException(mErrorComment)
            End If
            i = i + 1
            Call WriteLog(False, "RTC設定 ﾘﾄﾗｲ待ち　5秒")
            Call WaitStart(5)
            DoEvents()

        Loop

        '電源OFF
        Call I91(iName)
        PwrOFFWait = 30
        Call WriteLog(False, "電源遮断 " & PwrOFFWait & "秒待ち")
        Call WaitStart(PwrOFFWait)

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
                Throw New ApplicationException(mErrorComment)
            End If

            j = j + 1
            Call WriteLog(False, "2回目読込開始待ち・・・2秒")

            Call WaitStart(2)
            DoEvents()

        Loop

        'ResultData.RTC_Read = RecvDataLE
        Call WriteLog(False, "RTC読込 OK : " & RTCTime & " PCTime:" & pctime)

    End Sub

    Public Sub I18(ByVal iName As String)
        '==============================================================================================================
        ' Function : I18:RTC時刻ｶｳﾝﾄﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==============================================================================================================
        Dim RecvDataLE As String = ""
       
        'RTC確認 -------------------------------------
        Call WriteLog(False, "RTC確認")
        Dim RTCTime As String
        Dim BPsec1, BPsec2 As DateTime
        Dim PCsec1, PCsec2 As DateTime

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
            For i = 17 To 48 Step 2
                RTCTime = RTCTime & Chr(CInt("&H" & Mid(RecvDataLE, i, 2)))
            Next i

            If j = 0 Then
                BPsec1 = CDate(Mid(RTCTime, 9, 8)) : Call WriteLog(False, RTCTime) : j += 1
                PCsec1 = DateTime.Now
            Else
                BPsec2 = CDate(Mid(RTCTime, 9, 8))
                PCsec2 = DateTime.Now
                Select Case (BPsec2 - BPsec1).Seconds
                    Case 0
                        If (PCsec2 - PCsec1).Seconds > 3 Then
                            mErrorComment = "RTC時刻ｶｳﾝﾄ異常"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case 1, 2
                        If j = (BPsec2 - BPsec1).Seconds Then
                            Debug.Print(j.ToString)
                            Call WriteLog(False, RTCTime)
                            j += 1
                        End If
                    Case 3
                        Call WriteLog(False, RTCTime)
                        Call WriteLog(False, "3秒ｶｳﾝﾄ確認OK") : Exit Do
                End Select

            End If

            DoEvents()

        Loop

        Call WriteLog(False, "RTC時刻ｶｳﾝﾄ確認OK")

    End Sub

    Public Sub I19(ByVal iName As String)
        '==============================================================================================================
        ' Function : I19:電源電圧測定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I19(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            mTestMeasure = cSPPF.TestMeasure 'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I20(ByVal iName As String)
        '==============================================================================================================
        ' Function : I20:工場出荷設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I20(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub



    Public Sub I21(ByVal iName As String)
        '==============================================================================================================
        ' Function : I21:SIO最高速ﾎﾞｰﾚｰﾄ通信確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I21(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I22(ByVal iName As String)
        '==============================================================================================================
        ' Function : I22:IOﾊﾞｽ-XLFAILﾃｽﾄ(TEST MODE)
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I22(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I23(ByVal iName As String)
        '==============================================================================================================
        ' Function : I23:SIO最高速ﾎﾞｰﾚｰﾄ通信確認Mｺﾏﾝﾄﾞ
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
            If cSP.ComMode_Change("PF") = False Then
                mErrorComment = "Mｺﾏﾝﾄﾞ(0401)SIO最高速通信速度変更異常" : Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, cSP.LOG)

            Call WaitStart(1)   '待ち

            Call WriteLog(False, "Mｺﾏﾝﾄﾞ(0402)SIO最高速通信速度確認")
            Call WriteLog(False, "PF:19200bpsﾊﾟﾘﾃｨなし通信確認")
            If cSP.ComMode_Set("PF") = False Then
                mErrorComment = "Mｺﾏﾝﾄﾞ(0401)SIO最高速通信速度変更異常" : Throw New ApplicationException(mErrorComment)
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

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I50(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I51(ByVal iName As String)
        '==============================================================================================================
        ' Function : I51:IOﾊﾞｽ-XACSﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I51(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I52(ByVal iName As String)
        '==============================================================================================================
        ' Function : I52:BCLR TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I52(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I53(ByVal iName As String)
        '==============================================================================================================
        ' Function : I53:IOﾊﾞｽ-XBRﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I53(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I54(ByVal iName As String)
        '==============================================================================================================
        ' Function : I54:I/O BUS WRITE/READ TEST
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Dim SendParam As String() = {""}
        Dim RecvDATA As String()
        Dim mFuncName As String = ""
        Dim mSetDATA As String = ""

        Try
            cSP = New cls121_SP_TEST_MODE
            cSP.PORT = rs232cSPPort
            cSP.RS_COMM = Insfrm.SerialPort1

            For i = 0 To 2
                Select Case i
                    Case 0
                        mSetDATA = "AAAA0000" : mFuncName = "AAAA"
                    Case 1
                        mSetDATA = "55550000" : mFuncName = "5555"
                    Case 2
                        mSetDATA = "00000000" : mFuncName = "0000"
                End Select

                'ｽﾛｯﾄ13のYDのｲﾆｼｬﾗｲｽﾞ
                cSP.ICMD = 7 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 3
                ReDim SendParam(3)
                SendParam(1) = "700678" : SendParam(2) = "2" : SendParam(3) = "00000000" : cSP.SendPARA = SendParam
                Call cSP.TestComSendRecv()
                Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)

                'ｽﾛｯﾄ12のXDのｲﾆｼｬﾗｲｽﾞ
                cSP.ICMD = 7 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 3
                ReDim SendParam(3)
                SendParam(1) = "7005F8" : SendParam(2) = "2" : SendParam(3) = "00000000" : cSP.SendPARA = SendParam
                Call cSP.TestComSendRecv()
                Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)

                Call WriteLog(False, "ｽﾛｯﾄ13のYDに" & mFuncName & "をｾｯﾄ")
                cSP.ICMD = 7 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 3
                ReDim SendParam(3)
                SendParam(1) = "780660" : SendParam(2) = "4" : SendParam(3) = mSetDATA : cSP.SendPARA = SendParam
                Call cSP.TestComSendRecv()
                Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)

                If i = 2 Then Exit For

                Call WriteLog(False, "ｽﾛｯﾄ12のXDをﾘｰﾄﾞ")
                cSP.ICMD = 6 : cSP.IUNIT = 0 : cSP.ISLOT = 0 : cSP.IPARA = 2
                ReDim SendParam(2)
                SendParam(1) = "7805C0" : SendParam(2) = "4" : cSP.SendPARA = SendParam
                Call cSP.TestComSendRecv()
                Call WriteLog(False, "Send : " & cSP.SEND_TESTCOMMAND)
                Call WriteLog(False, "Recv : " & cSP.RECEIVE_DATA)
                RecvDATA = Split(cSP.RECEIVE_DATA, ",")
                If RecvDATA(1) <> "6" Or RecvDATA(5) <> mSetDATA Then
                    'mErrorComment = mFuncName & "異常"                                          'Rev0.20DEL
                    mErrorComment = mFuncName & "異常。" & RecvDATA(1) & ", " & RecvDATA(5)      'Rev0.20add
                    mTestSpec = "6, " & mSetDATA                         'Rev0.20add
                    Throw New ApplicationException(mErrorComment)
                Else
                    Call WriteLog(False, mFuncName & "OK")
                End If

            Next

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = mFuncName & "異常:" & ex.ToString
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP = Nothing
        End Try

        Call WriteLog(False, "I/O BUS WRITE/READ TEST OK")

    End Sub

    Public Sub I55(ByVal iName As String)
        '==============================================================================================================
        ' Function : I55:IOﾊﾞｽ-割込みﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I55(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I56(ByVal iName As String)
        '==============================================================================================================
        ' Function : I56:IOﾊﾞｽ-XLFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I56(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I57(ByVal iName As String)
        '==============================================================================================================
        ' Function : I57:IOﾊﾞｽ-XPFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I57(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I58(ByVal iName As String)
        '==============================================================================================================
        ' Function : I58:IOﾊﾞｽ-XSFAILﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I58(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I59(ByVal iName As String)
        '==============================================================================================================
        ' Function : I59:IOﾊﾞｽ検査前準備
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I59(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I90(ByVal iName As String)
        '==============================================================================================================
        ' Function : I90:ﾕﾆｯﾄ電源ON
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I90(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I91(ByVal iName As String)
        '==============================================================================================================
        ' Function : I91:ﾕﾆｯﾄ電源OFF
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I91(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I92(ByVal iName As String)
        '==============================================================================================================
        ' Function : I92:検査開始時処理
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I92(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
        End Try

    End Sub

    Public Sub I93(ByVal iName As String)
        '==============================================================================================================
        ' Function : I93:検査終了時メッセージ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.I93(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
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
    Public Sub IP1(ByVal iName As String)
        '==============================================================================================================
        ' Function : IP1:PrintLabel
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        cSPPF = New cls2D4_F3SPxx_PF_Inspection
        Try
            Call cSPPF.IP1(iName)
        Catch ex As Exception
            mErrorComment = cSPPF.ERR_COMMENT
            mTestSpec = cSPPF.TestSpec     'Rev0.20add
            Throw New ApplicationException(mErrorComment)
        Finally
            cSPPF = Nothing
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

        dbName = mAppPath & "210 Result Data\F3RP6XData.mdb"
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

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
    End Sub

    Private Function FirmRevRead() As String

        '検査型式の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "ProductSoftRev"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim Model As String = ""

        FirmRevRead = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName

            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            FirmRevRead = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try


    End Function

End Class
