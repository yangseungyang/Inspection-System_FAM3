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

Public Class cls2D1_F3SP6x_Inspection
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
    Dim cKP As cls124_KP_IOBUS_TEST

    Dim dfrm As frm500_MessagePopUp

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String

    Dim rs232cSPPort As String = "COM4"      ' "COM1" 
    'Dim rs232cCPU1SPPort As String = "COM4"

    Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3のLE
    Dim IP_U2 As String = "192.168.1.11"        'ﾒｲﾝCPU
    Dim IP_SP As String = "192.168.0.2"         '被検査用SP
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
            Case "I01" : Call I01(iName)            'I01:TestProDL-BootMode移行
            Case "I02" : Call I02(iName)            'I02:ﾒﾓﾘ検査-ﾌﾗｯｼｭ
            Case "I03" : Call I03(iName)            'I03:ﾒﾓﾘ検査-SDRAM
            Case "I04" : Call I04(iName)            'I04:ﾒﾓﾘ検査-共有SRAM
            Case "I05" : Call I05(iName)            'I05:TestMode移行
            Case "I06" : Call I06(iName)            'I06:TestMode終了
            Case "I07" : Call I07(iName)            'I07:CPU周辺-ﾓｼﾞｭｰﾙID確認
            Case "I08" : Call I08(iName)            'I08:CPU周辺-ﾊｰﾄﾞ種別確認
            Case "I09" : Call I09(iName)            'I09:CPU周辺-ROM SUM確認
            Case "I10" : Call I10(iName)            'I10:ｸﾛｯｸ検査
            Case "I11" : Call I11(iName)            'I11:SPU2周辺-SPU2
            Case "I12" : Call I12(iName)            'I12:SPU2周辺-DBﾗｯﾁ
            Case "I13" : Call I13(iName)            'I13:SPU2周辺-SPU/KP間割込み
            Case "I14" : Call I14(iName)            'I14:SPU2周辺-DMA RQ信号
            Case "I15" : Call I15(iName)            'I15:SPU2周辺-SCB用信号
            Case "I16" : Call I16(iName)            'I16:SPU2周辺-XWAIT信号
            Case "I17" : Call I17(iName)            'I17:SPU2周辺-ﾒﾓﾘ
            Case "I18" : Call I18(iName)            'I18:KP/FPGA/CPU検査-ﾊﾞｽｴﾗｰ
            Case "I19" : Call I19(iName)            'I19:KP/FPGA/CPU検査-ﾎﾙﾄ
            Case "I20" : Call I20(iName)            'I20:KP/FPGA-ﾊﾞｽｱﾋﾞ
            Case "I21" : Call I21(iName)            'I21:RTC検査
            Case "I22" : Call I22(iName)            'I22:LED検査
            Case "I23" : Call I23(iName)            'I23:USB検査-ﾌｧｰﾑ書き込み
            Case "I24" : Call I24(iName)            'I24:SIO-制御線
            Case "I25" : Call I25(iName)            'I25:SIO-ﾎﾞｰﾚｰﾄ
            Case "I26" : Call I26(iName)            'I26:ﾗﾀﾞｰ検査
            Case "I27" : Call I27(iName)            'I27:SDｶｰﾄﾞ検査-ﾗｲﾄﾌﾟﾛﾃｸﾄ
            Case "I28" : Call I28(iName)            'I28:SDｶｰﾄﾞ検査-ｶｰﾄﾞ電源/過電流保護
            Case "I29" : Call I29(iName)            'I29:Ethernet通信検査
            Case "I30" : Call I30(iName)            'I30:SW検査
            Case "I31" : Call I31(iName)            'I31:電源電圧検査
            Case "I32" : Call I32(iName)            'I32:ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟ検査
            Case "I33" : Call I33(iName)            'I33:EthernetLED検査
            Case "I34" : Call I34(iName)            'I34:ｺﾏﾝﾄﾞ割込み検査
            Case "I35" : Call I35(iName)            'I35:RTC設定
            Case "I36" : Call I36(iName)            'I36:ﾃﾞｰﾀ書込み
                'Case "I37" : Call I37(iName)            'I37:終了処理
            Case "I38" : Call I38(iName)            'I38:SDｶｰﾄﾞ検査-ﾉﾝﾗｲﾄﾌﾟﾛﾃｸﾄ

            Case "I40" : Call I40(iName)            'I40:SP69ﾌｧｰﾑ書込み
                'Case "I41" : Call I41(iName)            'I41:電源ON
            Case "I42" : Call I42(iName)            'I42:SP69形名照合
            Case "I43" : Call I43(iName)            'I43:主銘板ﾃﾞｰﾀ取得
            Case "I44" : Call I44(iName)            'I44:工場出荷設定
            Case "I45" : Call I45(iName)            'I45:ﾛｰﾀﾘｰｽｲｯﾁ0設定確認
            Case "I46" : Call I46(iName)            'I46:ﾛｰﾀﾘｰｽｲｯﾁC設定確認


            Case "I50" : Call I50(iName)            'I50:SLOT/XDFAIL TEST
            Case "I51" : Call I51(iName)            'I51:ACS SIGNAL TESTｽﾛｯﾄ5
            Case "I52" : Call I52(iName)            'I52:BCLR TESTｽﾛｯﾄ5
            Case "I53" : Call I53(iName)            'I53:BR TESTｽﾛｯﾄ5
                'Case "I54" : Call I54(iName)            'I54:I/O BUS WRITE/READ TEST
            Case "I55" : Call I55(iName)            'I55:I/O BUS INTERRUPT TESTｽﾛｯﾄ5
            Case "I56" : Call I56(iName)            'I56:XLFAIL TESTｽﾛｯﾄ5
            Case "I57" : Call I57(iName)            'I57:XPFAIL TESTｽﾛｯﾄ5
            Case "I58" : Call I58(iName)            'I58:XSFAIL TESTｽﾛｯﾄ5
            Case "I59" : Call I59(iName)            'I59:IOﾊﾞｽ検査前準備

            Case "I60" : Call I60(iName)            'I60:XIRQ4,7ｱｻｰﾄ確認ﾃｽﾄ

            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF
            Case "I92" : Call I92(iName)            'I92:検査開始時処理
            Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save        'Rev 0.26 KSJ ADD

            Case "IA2" : Call IA2(iName)            'IA2:MACｱﾄﾞﾚｽ重複確認

            Case "IP1" : Call IP1(iName)            'IP1:PrintLabel
            Case "IP2" : Call IP2(iName)            'IP2:OK PrintLabel

                'Case Else : mErrorComment = "検査プロシージャが定義されていません。"   'Rev0.11DEL
            Case Else : mErrorComment = "검사 프로시져가 정의 되어있지 않습니다."   'Rev0.11ADD
                MsgBox(mErrorComment)

        End Select


    End Sub

    Public Sub I01(ByVal iName As String)
        '==============================================================================================================
        ' Function : I01:TestProDL-BootMode移行
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rRes As String = ""
        Dim OpenFileName As String = mAppPath & "600 Equipment Program\SP6X\TestProgram\f3sp67tm.brf"
        Dim dfrm As New frm501_StatusDisp

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Call WriteLog(False, "")
        'Call WriteLog(True, iName)
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        Call WriteLog(False, "TestProDL-BootMode移行中・・・")

        'SP BootMode COM OPEN
        Try
            cISPComm = New cls101_ComRS232C
            cISPComm.RS_COMM = Insfrm.SerialPort1
            cISPComm.PORT = rs232cSPPort
            cISPComm.SETTING = "38400,N,8,1"
            Call cISPComm.PortSetup()
            cISPComm.TERMINAL_CHAR = vbCrLf
        Catch ex As Exception
            mErrorComment = "SPのCOMﾎﾟｰﾄ初期化に失敗しました。"
            cISPComm.PortClose()
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "SP BootMode COM OPEN OK!")
        '転送テストプログラムファイルの読込み
        Dim fs As New System.IO.FileStream(OpenFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'ファイルを読み込むバイト型配列を作成する
        Dim f(CInt(fs.Length) - 1) As Byte
        Dim g() As Byte
        'ファイルの内容をすべて読み込む
        fs.Read(f, 0, f.Length)
        '閉じる
        fs.Close()
        fs = Nothing

        Call WriteLog(False, "ﾃｽﾄﾌﾟﾛｸﾞﾗﾑ転送開始")

        'ファイル転送 -----------------------------------
        Dim j As Integer = 0 : Dim lenLFCount As Integer = 1
        Dim l As Integer = f.Length
        Dim strsend As String = ""
        Dim rcv As String = ""

        dfrm.Text = "ﾃｽﾄﾌﾟﾛｸﾞﾗﾑ転送中"

        Try
            Do While l >= lenLFCount '(最大ﾊﾞｲﾄ数で止め)

                Do 'LFまでのﾊﾞｲﾄ数(LF検知又は、最大ﾊﾞｲﾄ数で止め)
                    If lenLFCount = l Then Exit Do
                    If f(lenLFCount) = 10 Then Exit Do
                    lenLFCount = lenLFCount + 1
                Loop

                If lenLFCount = l Then
                    '送信ﾊﾞｲﾄ数(1行あたり)の決定
                    ReDim g(0 To lenLFCount - (j - 1) - 2)
                    '送信データの作成
                    For i = 1 To lenLFCount - (j - 1)
                        g(i - 1) = f(j)
                        strsend = strsend + Chr(g(i - 1))
                        j = j + 1
                        If i > UBound(g) Then
                            Exit For
                        End If
                    Next

                Else
                    '送信ﾊﾞｲﾄ数(1行あたり)の決定
                    ReDim g(0 To lenLFCount - (j - 1) - 1)
                    '送信データの作成
                    For i = 1 To lenLFCount - (j - 1)
                        g(i - 1) = f(j)
                        strsend = strsend + Chr(g(i - 1))
                        j = j + 1
                    Next
                End If

                Dim m As Integer = 0
                Do
                    '送信
                    Try
                        'frm020_TestInformation.MSComm1.Output = g()
                        cISPComm.SendByteData = g
                        Call cISPComm.FuncSendByteData()

                        '受信
                        If lenLFCount = l Then
                            Call cISPComm.FuncKYOUSEIRecvCmd()
                        Else
                            Call cISPComm.FuncRecvCmd()
                        End If

                        rcv = cISPComm.RECEIVE_DATA

                    Catch ex As Exception
                        'Call cISPComm.PortClose()
                        'cISPComm = Nothing
                        'mErrorComment = "ｼﾘｱﾙ送受信ｴﾗｰ"
                        'dfrm.Dispose()
                        'dfrm = Nothing
                        Call WriteLog(False, mErrorComment)
                        Throw New ApplicationException(mErrorComment)
                    End Try

                    If lenLFCount = l Then Exit Do '最終データはCR/LFが無い
                    If rcv <> "" Then Exit Do

                    If m >= 1 Then
                        mErrorComment = "返信がありません。"
                        Throw New ApplicationException(mErrorComment)
                    End If

                    Call WriteLog(False, "")
                    Call WriteLog(False, "===== Bootﾌﾟﾛｸﾞﾗﾑ送信 ﾘﾄﾗｲ待ち　5秒 =====")
                    DoEvents()

                    Call WriteLog(False, "Bootﾌﾟﾛｸﾞﾗﾑ送信 ﾘﾄﾗｲ待ち　5秒")
                    Call WaitTimer(5)

                    m = m + 1

                    DoEvents()

                Loop

                '受信データの表示
                dfrm.TextBox1.Text = "Send:" & g.Length.ToString & vbCrLf
                dfrm.TextBox1.Text = dfrm.TextBox1.Text & "Recv:" & rcv
                dfrm.ProgressBar1.Value = CInt((lenLFCount / l) * 100)

                Call dfrm.Show()
                If lenLFCount = l Then Exit Do '最終データはCR/LFが無い

                '初期化
                strsend = "" : rcv = ""

                lenLFCount = j

                '中止？
                'If .StopDownLoad = 1 Then
                '    ErrMSG = "中止が、押されました。"
                '    Throw New ApplicationException(mErrorComment)
                'End If

                DoEvents()
            Loop

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "ﾃｽﾄﾌﾟﾛｸﾞﾗﾑ転送異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            dfrm.Dispose()
            dfrm = Nothing
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "ﾃｽﾄﾌﾟﾛｸﾞﾗﾑDL 完了")
        Call WriteLog(False, "")

        Call WaitTimer(1)

        'ﾌﾞｰﾄﾓｰﾄﾞ ---------------------------------
        Call WriteLog(False, "ﾌﾞｰﾄｽﾄﾗｯﾌﾟﾓｰﾄﾞ確立中・・・")
        '送受信
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,N,8,1"
        cISPComm.TERMINAL_CHAR = vbCrLf

        Try
            cISPComm.SendCmd = "A"
            Call cISPComm.PortSetup()
            Call cISPComm.FuncSendCmd()

            Call WaitTimer(1)
            Call cISPComm.FuncKYOUSEIRecvCmd()

            rcv = cISPComm.RECEIVE_DATA
            rcv = Strings.Left(rcv, 1)
            If rcv <> ":" Then
                mErrorComment = "aｺﾏﾝﾄﾞ通信　異常"
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "ﾌﾞｰﾄｽﾄﾗｯﾌﾟﾓｰﾄﾞ 完了")

    End Sub

    Public Sub I02(ByVal iName As String)
        '==============================================================================================================
        ' Function : I02:ﾒﾓﾘ検査-ﾌﾗｯｼｭ
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
        Call WriteLog(False, "ﾒﾓﾘ検査-ﾌﾗｯｼｭ検査中・・・")

        Dim rcv As String

        'ﾏｰﾁﾝｸﾞ -------------------------
        Call WriteLog(False, "ﾌﾗｯｼｭ検査（ﾏｰﾁﾝｸﾞ）検査中・・・約１分待ち")
        Try
            Call CountStart(90)
            rcv = CommandComm("TS,24,0,0,7,0,0,", "Boot", 90)
            Call CountStop()
            If rcv <> "TS,24,0,0,0,7,107FFFFC,0,0," Then
                mErrorComment = "ﾌﾗｯｼｭ検査（ﾏｰﾁﾝｸﾞ）　異常　" & rcv
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "ﾌﾗｯｼｭ検査（ﾏｰﾁﾝｸﾞ） OK " & rcv)
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        'ﾊﾞｽ判定 ---------------------------
        Call WriteLog(False, "ﾌﾗｯｼｭ検査（ﾊﾞｽ判定）検査中・・・約６秒待ち")
        Try
            Call CountStart(10)
            rcv = CommandComm("TS,24,0,0,5,0,0,", "Boot", 15)
            Call CountStop()
            If rcv <> "TS,24,0,0,0,5,10000104,80000000,80000000," Then
                mErrorComment = "ﾌﾗｯｼｭ検査（ﾊﾞｽ判定）　異常　" & rcv
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "ﾌﾗｯｼｭ検査（ﾊﾞｽ判定） OK " & rcv)
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        'ﾁｯﾌﾟｲﾚｰｽ ---------------------------
        Call WriteLog(False, "ﾌﾗｯｼｭ検査（ﾁｯﾌﾟｲﾚｰｽ）検査中・・・約２４秒待ち")
        Try
            Call CountStart(60)
            rcv = CommandComm("TS,24,0,0,6,0,0,", "Boot", 60)
            Call CountStop()
            If rcv <> "TS,24,0,0,0,6,10000000,FFFFFFFF,FFFFFFFF," Then
                mErrorComment = "ﾌﾗｯｼｭ検査（ﾁｯﾌﾟｲﾚｰｽ）　異常　" & rcv
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "ﾌﾗｯｼｭ検査（ﾁｯﾌﾟｲﾚｰｽ） OK " & rcv)
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I03(ByVal iName As String)
        '==============================================================================================================
        ' Function : I03:ﾒﾓﾘ検査-SDRAM
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Dim ret As String = ""

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "ﾒﾓﾘ検査-SDRAM検査中・・・")

        Dim rcv As String

        'SRAM -------------------------
        Call WriteLog(False, "SDRAM検査中・・・約３６秒待ち")

        Try
            Call CountStart(60)
            rcv = CommandComm("TS,3,0,0,0,08000000,08FFFFFF,", "Boot", 40)
            Call CountStop()
            If rcv <> "TS,3,0,0,0,0,8FFFFFC,0,0," Then
                mErrorComment = "SDRAM検査　異常　" & rcv
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "SDRAM検査 OK " & rcv)
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I04(ByVal iName As String)
        '==============================================================================================================
        ' Function : I04:ﾒﾓﾘ検査-共有SRAM
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "ﾒﾓﾘ検査-共有SRAM検査中・・・")

        Dim rcv As String

        'SRAM -------------------------
        Call WriteLog(False, "共有SRAM検査中・・・約３秒待ち")

        Try
            rcv = CommandComm("TS,3,0,0,0,16800000,1687FFFF,", "Boot", 5)
            If rcv <> "TS,3,0,0,0,0,1687FFFC,0,0," Then
                mErrorComment = "共有SRAM検査　異常　" & rcv
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "共有SRAM検査 OK " & rcv)

        Catch ex As Exception
            mErrorComment = "共有SRAM検査　異常　"
            Throw New ApplicationException(mErrorComment)
        End Try

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I05(ByVal iName As String)
        '==============================================================================================================
        ' Function : I05:TestMode移行
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        Call WriteLog(False, "TestMode移行中・・・")
        Try
            If TestModeStart() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "ﾃｽﾄﾓｰﾄﾞ移行　完了")

        ' =============================================================

    End Sub

    Public Sub I06(ByVal iName As String)
        '==============================================================================================================
        ' Function : I06:TestMode終了
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rcv As String
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        Call WriteLog(False, "TestMode終了中・・・")

        Try
            rcv = CommandComm("TS,2,0,0,", "Test", 5)
            If rcv <> "TS,2,0,0,0," Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "TestMode終了　異常"
            Debug.Print(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "TestMode終了OK")

    End Sub

    Public Sub I07(ByVal iName As String)
        '==============================================================================================================
        ' Function : I07:CPU周辺-ﾓｼﾞｭｰﾙID確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rcv As String
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        Call WriteLog(False, "CPU周辺-ﾓｼﾞｭｰﾙID確認検査中・・・")

        '検査型式の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "MSCODE"
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

            Model = sData(0)

        Catch ex As Exception

        Finally
            cADO = Nothing
        End Try

        Try
            rcv = CommandComm("TS,13,0,0,", "Test", 5)
            Select Case Mid(Model, 3, 4)                    'Ver0.01にて追加
                Case "SP69"
                    If Mid(rcv, 13, 4) <> "SP67" Then       'SP69は機能検査時はSP67で検査するため
                        mErrorComment = "ﾓｼﾞｭｰﾙID検査　異常 " & rcv
                        Throw New ApplicationException(mErrorComment)
                    End If
                Case Else
                    If Mid(rcv, 13, 4) <> Mid(Model, 3, 4) Then
                        mErrorComment = "ﾓｼﾞｭｰﾙID検査　異常 " & rcv
                        Throw New ApplicationException(mErrorComment)
                    End If
            End Select

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "ﾓｼﾞｭｰﾙID検査　異常"
            Call WriteLog(False, mErrorComment)
            Debug.Print(mErrorComment)
            Throw New ApplicationException(mErrorComment)

        End Try

        Call WriteLog(False, "CPU周辺-ﾓｼﾞｭｰﾙID確認検査OK")

    End Sub

    Public Sub I08(ByVal iName As String)
        '==============================================================================================================
        ' Function : I08:CPU周辺-ﾊｰﾄﾞ種別確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rcv As String
        Dim Model As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        '検査型式の読込み
        Model = ModelRead()

        'Test
        Call WriteLog(False, "CPU周辺-ﾊｰﾄﾞ種別確認中・・・")
        Try
            rcv = CommandComm("TS,6,0,0,0021C324,4,", "Test", 5)

            If InStr(Model, "F3SP66") <> 0 Then         'Mid(Model, 1, 6) = "F3SP66" Then
                '        If rcv <> "TS,6,0,0,0,4,FFFF67C0,21C324," Then
                If Mid(rcv, 18, 4) <> "67C0" Then
                    mErrorComment = "ﾊｰﾄﾞ種別確認　異常 " & rcv
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If
            Else
                '        If rcv <> "TS,6,0,0,0,4,FFFF77C0,21C324," Then
                If Mid(rcv, 18, 4) <> "77C0" Then
                    mErrorComment = "ﾊｰﾄﾞ種別確認　異常 " & rcv
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If
            End If

        Catch ex As Exception
            mErrorComment = "ﾊｰﾄﾞ種別確認　異常 "
            Debug.Print(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally

        End Try

        Call WriteLog(False, "CPU周辺-ﾊｰﾄﾞ種別確認 OK")

    End Sub

    Public Sub I09(ByVal iName As String)
        '==============================================================================================================
        ' Function : I09:CPU周辺-ROM SUM確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rcv As String
        Dim Model As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        '検査型式の読込み
        Model = ModelRead()

        Dim dbName As String
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim ROM_SUM66 As String = ""
        Dim ROM_SUM67 As String = ""
        Dim ROM_SUMBT As String = ""

        'チェックサムの読込み
        dbName = mAppPath & "520 Product Firmware\SP6X\FirmName.mdb"
        tbName = "tblCheckSUM"
        rField = "CheckSUM66,CheckSUM67,CheckSUMBT"
        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            cADO.TB_NAME = tbName
            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)
            ROM_SUM66 = sData(0) : ROM_SUM67 = sData(1) : ROM_SUMBT = sData(2)
        Catch ex As Exception
            cADO = Nothing
        End Try

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "CPU周辺-ROM SUM確認中・・・")

        'CPU周辺-ROM SUM(Firm)確認 -------------------------
        WriteLog(False, "CPU周辺-ROM SUM(Firm)確認中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,4,0,0,0,0,", "Test", 5)
            If InStr(Model, "F3SP66") <> 0 Then         'Mid(Model, 1, 6) = "F3SP66" Then
                'ResultData.ROM_SUM66 = Mid(rcv, 12, Len(rcv) - 12)
                If rcv <> "TS,4,0,0,0," & ROM_SUM66 & "," Then
                    mErrorComment = "CPU周辺-ROM SUM(Firm)確認　異常 " & rcv
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If
                Call WriteLog(False, "CPU周辺-ROM SUM(Firm)確認 OK " & rcv)
            Else
                'ResultData.ROM_SUM67 = Mid(rcv, 12, Len(rcv) - 12)
                If rcv <> "TS,4,0,0,0," & ROM_SUM67 & "," Then
                    mErrorComment = "CPU周辺-ROM SUM(Firm)確認　異常 " & rcv
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If
                Call WriteLog(False, "CPU周辺-ROM SUM(Firm)確認 OK " & rcv)
            End If

        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try


        'CPU周辺-ROM SUM(Boot)確認 -------------------------
        Call WriteLog(False, "CPU周辺-ROM SUM(BOOT)確認中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,29,0,0,", "Test", 5)
            If rcv <> "TS,29,0,0,0," & ROM_SUMBT & "," Then
                mErrorComment = "CPU周辺-ROM SUM(BOOT)確認　異常 " & rcv
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "CPU周辺-ROM SUM(BOOT)確認 OK " & rcv)

    End Sub

    Public Sub I10(ByVal iName As String)
        '==============================================================================================================
        ' Function : I10:ｸﾛｯｸ検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Dim RcvCmd As String = ""

        '検査規格定数宣言
        Dim TestStd_CPUCoreClockH As String = "100"
        Dim TestStd_CPUCoreClockL As String = "70"
        Dim TestStd_CPUBusClockH As String = "100"
        Dim TestStd_CPUBusClockL As String = "70"
        Dim TestStd_KPInnerClockH As String = "2A00"
        Dim TestStd_KPInnerClockL As String = "2800"
        Dim TestStd_KPBusClockH As String = "9200"
        Dim TestStd_KPBusClockL As String = "8F00"
        Dim TestStd_SPU2ClockH As String = "7000"
        Dim TestStd_SPU2ClockL As String = "6F00"

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "ｸﾛｯｸ検査中・・・")

        Dim rcv As String
        Dim RcvData As String
        Dim FuncName As String
        Dim StdMax As String
        Dim StdMin As String

        'CPUｺｱ検査 -------------------------
        FuncName = "CPUｺｱ確認"
        StdMax = TestStd_CPUCoreClockH
        StdMin = TestStd_CPUCoreClockL

        Call WriteLog(False, FuncName & "中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,20,0,0,1,1000," & StdMin & "," & StdMax & ",", "Test", 5)
            RcvData = Mid(rcv, 15, Len(rcv) - 15)
            If Microsoft.VisualBasic.Left(rcv, 11) <> "TS,20,0,0,0" Then
                mErrorComment = FuncName & "測定 異常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, FuncName & "測定 正常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax)

        'CPUﾊﾞｽ検査 -------------------------
        FuncName = "CPUﾊﾞｽ確認"
        StdMax = TestStd_CPUBusClockH
        StdMin = TestStd_CPUBusClockL

        Call WriteLog(False, FuncName & "中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,20,0,0,2,1000," & StdMin & "," & StdMax & ",", "Test", 5)
            RcvData = Mid(rcv, 15, Len(rcv) - 15)
            If Microsoft.VisualBasic.Left(rcv, 11) <> "TS,20,0,0,0" Then
                mErrorComment = FuncName & "測定 異常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, FuncName & "測定 正常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax)

        'KP内部検査 -------------------------
        FuncName = "KP内部確認"
        StdMax = TestStd_KPInnerClockH
        StdMin = TestStd_KPInnerClockL

        Call WriteLog(False, FuncName & "中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,20,0,0,3,1000," & StdMin & "," & StdMax & ",", "Test", 5)
            RcvData = Mid(rcv, 15, Len(rcv) - 15)
            If Microsoft.VisualBasic.Left(rcv, 11) <> "TS,20,0,0,0" Then
                mErrorComment = FuncName & "測定 異常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, FuncName & "測定 正常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax)

        'KPﾊﾞｽ検査 -------------------------
        FuncName = "KPﾊﾞｽ確認"
        StdMax = TestStd_KPBusClockH
        StdMin = TestStd_KPBusClockL

        Call WriteLog(False, FuncName & "中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,20,0,0,4,1000," & StdMin & "," & StdMax & ",", "Test", 5)
            RcvData = Mid(rcv, 15, Len(rcv) - 15)
            If Microsoft.VisualBasic.Left(rcv, 11) <> "TS,20,0,0,0" Then
                mErrorComment = FuncName & "測定 異常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, FuncName & "測定 正常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax)

        'SPU2検査 -------------------------
        FuncName = "SPU2確認"
        StdMax = TestStd_SPU2ClockH
        StdMin = TestStd_SPU2ClockL

        Call WriteLog(False, FuncName & "中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,5,0,0,1," & StdMin & "," & StdMax & ",0,", "Test", 5)
            RcvData = Mid(rcv, 16, Len(rcv) - 16)
            If Microsoft.VisualBasic.Left(rcv, 10) <> "TS,5,0,0,0" Then
                mErrorComment = FuncName & "測定 異常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, FuncName & "測定 正常 測定値=" & RcvData & "　規格：" & StdMin & "～" & StdMax)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I11(ByVal iName As String)
        '==============================================================================================================
        ' Function : I11:SPU2周辺-SPU2
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mMeas As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "SPU2周辺-SPU2検査中・・・")

        Dim rcv As String

        'SPU2周辺-SPU2検査 -------------------------
        Call WriteLog(False, "SPU2周辺-SPU2検査中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,5,0,0,0,0,0,0,", "Test", 5)
            If rcv <> "TS,5,0,0,0,0,0,0," Then
                mErrorComment = "SPU2周辺-SPU2検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "SPU2周辺-SPU2検査 OK " & rcv)

    End Sub

    Public Sub I12(ByVal iName As String)
        '==============================================================================================================
        ' Function : I12:SPU2周辺-DBﾗｯﾁ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mMeas As String = ""

        Dim RcvCmd As String = ""
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "SPU2周辺-DBﾗｯﾁ検査中・・・")

        Dim rcv As String

        'SPU2周辺-DBﾗｯﾁ検査 -------------------------
        Call WriteLog(False, "SPU2周辺-DBﾗｯﾁ検査中・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,5,0,0,2,0,0,0,", "Test", 5)
            If rcv <> "TS,5,0,0,0,2,0,0," Then
                mErrorComment = "SPU2周辺-DBﾗｯﾁ検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "SPU2周辺-DBﾗｯﾁ検査 OK " & rcv)

    End Sub

    Public Sub I13(ByVal iName As String)
        '==============================================================================================================
        ' Function : I13:SPU2周辺-SPU/KP間割込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mMeas As String = ""

        Dim RcvCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "SPU2周辺-SPU/KP間割込検査中・・・")
        '
        '
        '
        '
        '
        '
        '
        '
        ' ---------------------------------------------------------------------------------------


    End Sub

    Public Sub I14(ByVal iName As String)
        '==============================================================================================================
        ' Function : I14:SPU2周辺-DMA RQ信号
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mMeas As String = ""

        Dim RcvCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "SPU2周辺-DMA RQ信号検査中・・・")
        '
        '
        '
        '
        '
        '
        '
        '
        ' ---------------------------------------------------------------------------------------


    End Sub

    Public Sub I15(ByVal iName As String)
        '==============================================================================================================
        ' Function : I15:SPU2周辺-SCB用信号
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mMeas As String = ""

        Dim RcvCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "SPU2周辺-SCB用信号検査中・・・")
        '
        '
        '
        '
        '
        '
        '
        '
        ' ---------------------------------------------------------------------------------------


    End Sub

    Public Sub I16(ByVal iName As String)
        '==============================================================================================================
        ' Function : I16:SPU2周辺-XWAIT信号
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "SPU2周辺-XWAIT信号検査中・・・")

        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = vbCrLf

        Dim rcv As String

        'ｻｲｸﾙﾓｰﾄﾞ変更 -------------------------
        Call WriteLog(False, "ｻｲｸﾙﾓｰﾄﾞ変更中・・・約２秒待ち")
        cISPComm.SendCmd = "TS,7,0,0,16F00002,2,0,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,7,0,0,0,2,16F00002," Then
                mErrorComment = "ｻｲｸﾙﾓｰﾄﾞ変更　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "ｻｲｸﾙﾓｰﾄﾞ変更 OK " & rcv)

        'SQﾒﾓﾘﾏｰﾁﾝｸﾞ -------------------------
        Call WriteLog(False, "SQﾒﾓﾘﾏｰﾁﾝｸﾞ検査中・・・約４秒待ち")
        cISPComm.SendCmd = "TS,3,0,0,0,16A00000,16AFFFFF,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,3,0,0,0,0,16AFFFFC,0,0," Then
                mErrorComment = "SQﾒﾓﾘﾏｰﾁﾝｸﾞ検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "SQﾒﾓﾘﾏｰﾁﾝｸﾞ検査 OK " & rcv)

        'ﾊﾞｰｽﾄﾓｰﾄﾞ変更 -------------------------
        Call WriteLog(False, "ﾊﾞｰｽﾄﾓｰﾄﾞ変更中・・・約２秒待ち")
        cISPComm.SendCmd = "TS,7,0,0,16F00002,2,1,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,7,0,0,0,2,16F00002," Then
                mErrorComment = "ﾊﾞｰｽﾄﾓｰﾄﾞ変更　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "ﾊﾞｰｽﾄﾓｰﾄﾞ変更 OK " & rcv)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I17(ByVal iName As String)
        '==============================================================================================================
        ' Function : I17:SPU2周辺-ﾒﾓﾘ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = vbCrLf

        Call WriteLog(False, "SPU周辺-ﾒﾓﾘ検査中・・・")

        Dim rcv As String

        'DVBﾒﾓﾘ指定 -------------------------
        Call WriteLog(False, "DVBﾒﾓﾘ指定中・・・約２秒待ち")
        cISPComm.SendCmd = "TS,7,0,0,16F00008,2,7,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,7,0,0,0,2,16F00008," Then
                mErrorComment = "DVBﾒﾓﾘ指定　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "DVBﾒﾓﾘ指定 OK " & rcv)

        'DVBメモリテスト -------------------------
        Call WriteLog(False, "DVBﾒﾓﾘ検査中・・・約４秒待ち")
        cISPComm.SendCmd = "TS,3,0,0,0,16C00000,16CFFFFF,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,3,0,0,0,0,16CFFFFC,0,0," Then
                mErrorComment = "DVBﾒﾓﾘ検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "DVBﾒﾓﾘ検査 OK " & rcv)

        'DVﾒﾓﾘ指定 -------------------------
        Call WriteLog(False, "DVﾒﾓﾘ指定中・・・約２秒待ち")
        cISPComm.SendCmd = "TS,7,0,0,16F00008,2,6,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,7,0,0,0,2,16F00008," Then
                mErrorComment = "DVﾒﾓﾘ指定　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "DVﾒﾓﾘ指定 OK " & rcv)

        'DVメモリテスト -------------------------
        Call WriteLog(False, "DVﾒﾓﾘ検査中・・・約４秒待ち")
        cISPComm.SendCmd = "TS,3,0,0,0,16C00000,16CFFFFF,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,3,0,0,0,0,16CFFFFC,0,0," Then
                mErrorComment = "DVﾒﾓﾘ検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "DVﾒﾓﾘ検査 OK " & rcv)

        'MCメモリテスト -------------------------
        Call WriteLog(False, "MCﾒﾓﾘ検査中・・・約３秒待ち")
        cISPComm.SendCmd = "TS,3,0,0,0,16E00000,16E7FFFF,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,3,0,0,0,0,16E7FFFC,0,0," Then
                mErrorComment = "MCﾒﾓﾘ検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "MCﾒﾓﾘ検査 OK " & rcv)


    End Sub

    Public Sub I18(ByVal iName As String)
        '==============================================================================================================
        ' Function : I18:KP/FPGA/CPU検査-ﾊﾞｽｴﾗｰ
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
        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = vbCrLf

        Call WriteLog(False, "KP/FPGA/CPU検査-ﾊﾞｽｴﾗｰ検査中・・・")

        Dim rcv As String

        'ﾊﾞｽｴﾗｰ検査-------------------------
        Call WriteLog(False, "ﾊﾞｽｴﾗｰ検査中・・・約２秒待ち")
        cISPComm.SendCmd = "TS,C,0,0,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,C,0,0,0," Then
                mErrorComment = "ﾊﾞｽｴﾗｰ検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "ﾊﾞｽｴﾗｰ検査 OK " & rcv)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I19(ByVal iName As String)
        '==============================================================================================================
        ' Function : I19:KP/FPGA/CPU検査-ﾎﾙﾄ
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
        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = vbCrLf

        Call WriteLog(False, "KP/FPGA/CPU検査-ﾎﾙﾄ検査中・・・")

        Dim rcv As String

        'ﾎﾙﾄ検査-------------------------
        Call WriteLog(False, "ﾎﾙﾄ検査中・・・約２秒待ち")
        cISPComm.SendCmd = "TS,26,0,0,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,26,0,0,0," Then
                mErrorComment = "ﾎﾙﾄ検査　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "ﾎﾙﾄ検査 OK " & rcv)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I20(ByVal iName As String)
        '==============================================================================================================
        ' Function : I20:KP/FPGA-ﾊﾞｽｱﾋﾞ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Dim rRes As String = ""
        Dim rCode As String = ""
        Dim rData() As String = {}

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        ' ---------------------------------------------------------------------------------------
        Dim Result As String
        Dim sModel As String
        Dim sChr As String
        Dim i As Integer
        Dim SigName As String
        Dim RecvLEData As String = ""

        'CPU1経由SP6ﾃﾞｰﾀ読込み -------------------------------------
        SigName = "CPU1経由SP6ﾃﾞｰﾀ読込み"
        Call WriteLog(False, SigName & "検査中・・・")
        Call WaitStart(5)
        'CPU1 SP SerialSetting 
        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U2
        Call cLnk.F3PCLinkCommunication("02WRDZ00121,03", RecvLEData)
        cLnk = Nothing
        Result = Mid(RecvLEData, 5, 12)

        sModel = ""
        For i = 1 To Len(Result) Step 2
            sChr = "&h" & (Mid(Result, i, 2))
            sModel = sModel & Chr(CInt(sChr))
        Next

        '検査型式の読込み
        Dim Model As String = ""
        Model = ModelRead()

        'Ver0.01にて追加        SP69は検査時はSP67として検査
        Select Case Strings.Left(Model, 6)
            Case "F3SP69"
                If sModel <> "F3SP67" Then
                    mErrorComment = SigName & "CPU1経由SP6ﾃﾞｰﾀ　異常:" & sModel
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If
            Case Else
                If Strings.Left(Model, 6) <> sModel Then
                    mErrorComment = SigName & "CPU1経由SP6ﾃﾞｰﾀ　異常:" & sModel
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If
        End Select

        Call WriteLog(False, "検査　正常　CPU1経由SP6ﾃﾞｰﾀ：" & sModel)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I21(ByVal iName As String)
        '==============================================================================================================
        ' Function : I21:RTC検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mMeas As String = ""
        Dim mSetting As String = ""
        Dim mDelim As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "RTC検査中・・・")

        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = vbCrLf

        Dim rcv As String

        'RTC検査-------------------------
        Call WriteLog(False, "RTC検査中・・・約３秒待ち")

        cISPComm.SendCmd = "TS,8,0,0,"
        Try
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,8,0,0,0,0," Then
                mErrorComment = "RTC検査　異常 " & rcv
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Throw New ApplicationException(mErrorComment)
        Finally
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "RTC検査 OK " & rcv)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I22(ByVal iName As String)
        '==============================================================================================================
        ' Function : I22:LED検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim rcv As String
        Dim Timing As Integer
        Dim fCam As Boolean = False
        Dim Model As String = ModelRead()

        Dim FuncName As String = ""
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        Dim cEI As cls130_EI_GAZO

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "LED検査中・・・")

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

        Select fCam
            Case True       'カメラありのとき
                Try
                    'SP TestMode COM OPEN
                    cISPComm = New cls101_ComRS232C
                    cISPComm.RS_COMM = Insfrm.SerialPort1
                    cISPComm.PORT = rs232cSPPort
                    cISPComm.SETTING = "9600,E,8,1"
                    cISPComm.TERMINAL_CHAR = vbCrLf
                    cISPComm.PortSetup()
                    cISPComm.SendCmd = "TS,27,0,0,"
                    Call cISPComm.FuncSendCmd()
                    Call cISPComm.FuncRecvCmd()
                    rcv = cISPComm.RECEIVE_DATA
                    If rcv <> "TS,27,0,0,0," Then
                        mErrorComment = "LED検査　異常 " & rcv
                        Call WriteLog(False, mErrorComment)
                        Throw New ApplicationException(mErrorComment)
                    End If
                    cISPComm.SendCmd = "TS,7,0,0,16200016,2,0," : Call cISPComm.FuncSendCmd()
                    cISPComm.SendCmd = "TS,7,0,0,1660B21B,1,0," : Call cISPComm.FuncSendCmd()

                    Call WriteLog(False, "LEDﾌｨﾙﾀﾁｪｯｸ")
                    Call mSecWait(500)
                    Call cEI.GAZO_TEST() : Call WriteLog(False, cEI.ResultData)
                    If cEI.RESULT <> "0" Then
                        'メッセージ表示
                        dfrm = New frm500_MessagePopUp
                        'dfrm.DISP_TITLENAME = "LEDﾌｨﾙﾀ検査"                                                 'Rev0.11DEL
                        'dfrm.DISP_MESSAGE = "LEDﾌｨﾙﾀｼﾙｸ印刷は合っていますか？MODEL=" & Mid(Model, 1, 9)     'Rev0.11DEL
                        dfrm.DISP_TITLENAME = "LED필터 검사"                                                 'Rev0.11ADD
                        'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까? MODEL=" & Mid(Model, 1, 9)      'Rev0.11ADD     Ver 0.23 JDI DEL
                        'dfrm.DISP_MESSAGE = "LED필터 실크인쇄가 정확합니까?" & vbCrLf & "MODEL=" & Mid(Model, 3, 7)         'Ver 0.23 JDI ADD /  Ver 0.31 GYS DEL
                        dfrm.DISP_MESSAGE = "필터 형명 확인" & vbCrLf & " = " & Mid(Model, 3, 7)                     'Ver 0.31 GYS ADD
                        dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                        Call dfrm.frmShow()

                        Do
                            'wait push button -------------------
                            If dfrm.DIALOG_RESULT = "NG" Then
                                dfrm.Close()
                                dfrm = Nothing
                                mErrorComment = "LEDﾌｨﾙﾀ検査検査 異常 "
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

                    cEI.MasterFile = "F3SP76-7N_LED.jpg"
                    Call cEI.MasterFileRead()

                    For i = 1 To 15
                        Select Case i
                            '-----< Rev0.11 DEL/ADD >-----
                            'Case 1 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,1660B21B,1,C0,"
                            'Case 2 : FuncName = "RDY,RUN点灯" : cISPComm.SendCmd = "TS,7,0,0,1660B228,1,80,"
                            'Case 3 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,1660B228,1,0,"
                            'Case 4 : FuncName = "RDY,ALM点灯" : cISPComm.SendCmd = "TS,7,0,0,1660B229,1,80,"
                            'Case 5 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,1660B229,1,0,"
                            'Case 6 : FuncName = "RDY,ERR点灯" : cISPComm.SendCmd = "TS,7,0,0,1660B22A,1,80,"
                            'Case 7 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,1660B22A,1,0,"
                            'Case 8 : FuncName = "RDY,SD点灯" : cISPComm.SendCmd = "TS,7,0,0,16200018,2,8000,"
                            'Case 9 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,16200018,2,0,"
                            'Case 10 : FuncName = "RDY,EXE点灯" : cISPComm.SendCmd = "TS,7,0,0,1620001A,2,8000,"
                            'Case 11 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,1620001A,2,0,"
                            'Case 12 : FuncName = "RDY,US1点灯" : cISPComm.SendCmd = "TS,7,0,0,1620001C,2,8000,"
                            'Case 13 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,1620001C,2,0,"
                            'Case 14 : FuncName = "RDY,US2点灯" : cISPComm.SendCmd = "TS,7,0,0,1620001E,2,8000,"
                            'Case 15 : FuncName = "RDY点灯" : cISPComm.SendCmd = "TS,7,0,0,1620001E,2,0,"
                            Case 1 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,1660B21B,1,C0,"
                            Case 2 : FuncName = "RDY,RUN점등" : cISPComm.SendCmd = "TS,7,0,0,1660B228,1,80,"
                            Case 3 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,1660B228,1,0,"
                            Case 4 : FuncName = "RDY,ALM점등" : cISPComm.SendCmd = "TS,7,0,0,1660B229,1,80,"
                            Case 5 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,1660B229,1,0,"
                            Case 6 : FuncName = "RDY,ERR점등" : cISPComm.SendCmd = "TS,7,0,0,1660B22A,1,80,"
                            Case 7 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,1660B22A,1,0,"
                            Case 8 : FuncName = "RDY,SD점등" : cISPComm.SendCmd = "TS,7,0,0,16200018,2,8000,"
                            Case 9 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,16200018,2,0,"
                            Case 10 : FuncName = "RDY,EXE점등" : cISPComm.SendCmd = "TS,7,0,0,1620001A,2,8000,"
                            Case 11 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,1620001A,2,0,"
                            Case 12 : FuncName = "RDY,US1점등" : cISPComm.SendCmd = "TS,7,0,0,1620001C,2,8000,"
                            Case 13 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,1620001C,2,0,"
                            Case 14 : FuncName = "RDY,US2점등" : cISPComm.SendCmd = "TS,7,0,0,1620001E,2,8000,"
                            Case 15 : FuncName = "RDY점등" : cISPComm.SendCmd = "TS,7,0,0,1620001E,2,0,"
                                '-------------------------------
                        End Select

                        Try
                            Call cISPComm.FuncSendCmd()
                            Call mSecWait(500)
                        Catch ex As Exception
                            Throw New ApplicationException(mErrorComment)
                        End Try

                        Call cEI.GAZO_TEST()

                        Try
                            Select Case FuncName
                                '-----< Rev0.11 DEL/ADD >-----
                                'Case "RDY点灯" : If cEI.RESULT <> "0,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RUN点灯" : If cEI.RESULT <> "1,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,ALM点灯" : If cEI.RESULT <> "0,1,0,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,ERR点灯" : If cEI.RESULT <> "0,1,1,0,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,SD点灯" : If cEI.RESULT <> "0,1,1,1,0,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,EXE点灯" : If cEI.RESULT <> "0,1,1,1,1,0,1,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,US1点灯" : If cEI.RESULT <> "0,1,1,1,1,1,0,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,US2点灯" : If cEI.RESULT <> "0,1,1,1,1,1,1,0,1,1,1,1" Then GoTo ErrCheck
                                'Case "全消灯" : If cEI.RESULT <> "1,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                'Case "RDY,RUN点灯" : If cEI.RESULT <> "0,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY점등" : If cEI.RESULT <> "0,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RUN점등" : If cEI.RESULT <> "1,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY,ALM점등" : If cEI.RESULT <> "0,1,0,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY,ERR점등" : If cEI.RESULT <> "0,1,1,0,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY,SD점등" : If cEI.RESULT <> "0,1,1,1,0,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY,EXE점등" : If cEI.RESULT <> "0,1,1,1,1,0,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY,US1점등" : If cEI.RESULT <> "0,1,1,1,1,1,0,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY,US2점등" : If cEI.RESULT <> "0,1,1,1,1,1,1,0,1,1,1,1" Then GoTo ErrCheck
                                Case "전소등" : If cEI.RESULT <> "1,1,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                Case "RDY,RUN점등" : If cEI.RESULT <> "0,0,1,1,1,1,1,1,1,1,1,1" Then GoTo ErrCheck
                                    '-------------------------------
                            End Select

                            GoTo NextGo

ErrCheck:
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
                    Throw New ApplicationException(mErrorComment)
                Finally
                    Call cISPComm.PortClose()
                    cISPComm = Nothing
                End Try

            Case False

                'LED検査-------------------------
                Try
                    'メッセージ表示
                    dfrm = New frm500_MessagePopUp
                    dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\LED.ppt"
                    'dfrm.DISP_TITLENAME = "LED 検査"                                  'Rev0.11DEL
                    dfrm.DISP_TITLENAME = "LED 검사"                                  'Rev0.11ADD
                    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

                    'SP TestMode COM OPEN
                    cISPComm = New cls101_ComRS232C
                    cISPComm.RS_COMM = Insfrm.SerialPort1
                    cISPComm.PORT = rs232cSPPort
                    cISPComm.SETTING = "9600,E,8,1"
                    cISPComm.TERMINAL_CHAR = vbCrLf
                    cISPComm.PortSetup()

                    Call dfrm.frmShow()

                    '検査開始
                    cISPComm.SendCmd = "TS,27,0,0,"
                    Try
                        Call WriteLog(False, cISPComm.SendCmd)
                        Call cISPComm.FuncSendCmd()
                        Call cISPComm.FuncRecvCmd()
                        Call WriteLog(False, cISPComm.RECEIVE_DATA)
                        rcv = cISPComm.RECEIVE_DATA
                        If rcv <> "TS,27,0,0,0," Then
                            mErrorComment = "LED検査　異常 " & rcv
                            Call WriteLog(False, mErrorComment)
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Catch ex As Exception
                        Call cISPComm.PortClose()
                        cISPComm = Nothing
                        Throw New ApplicationException(mErrorComment)
                    End Try

                    Timing = 500

                    '1消灯
                    'frm020_TestInformation.MSComm1.Output = "TS,7,0,0,16200016,2,0," & vbCrLf
                    cISPComm.SendCmd = "TS,7,0,0,16200016,2,0," : Call cISPComm.FuncSendCmd()
                    Call mSecWait(Timing)

                    'RDY消灯
                    'frm020_TestInformation.MSComm1.Output = "TS,7,0,0,1660B21B,1,0," & vbCrLf
                    cISPComm.SendCmd = "TS,7,0,0,1660B21B,1,0," : Call cISPComm.FuncSendCmd()
                    Call mSecWait(Timing)

                    Do
                        For i = 0 To 23
                            Select Case i
                                Case 0 : cISPComm.SendCmd = "TS,7,0,0,1660B21B,1,C0,"       'RDY点灯
                                Case 1 : cISPComm.SendCmd = "TS,7,0,0,1660B228,1,80,"       'RUN点灯
                                Case 2 : cISPComm.SendCmd = "TS,7,0,0,1660B228,1,0,"        'RUN消灯
                                Case 3 : cISPComm.SendCmd = "TS,7,0,0,1660B21B,1,0,"        'RDY消灯
                                Case 4 : cISPComm.SendCmd = "TS,7,0,0,1660B229,1,80,"       'ALM点灯
                                Case 5 : cISPComm.SendCmd = "TS,7,0,0,1660B229,1,0,"        'ALM消灯
                                Case 6 : cISPComm.SendCmd = "TS,7,0,0,1660B22A,1,80,"       'ERR点灯
                                Case 7 : cISPComm.SendCmd = "TS,7,0,0,1660B22A,1,0,"        'ERR消灯
                                Case 8 : cISPComm.SendCmd = "TS,7,0,0,16200018,2,8000,"     'SD点灯
                                Case 9 : cISPComm.SendCmd = "TS,7,0,0,16200018,2,0,"        'SD消灯
                                Case 10 : cISPComm.SendCmd = "TS,7,0,0,1620001A,2,8000,"    'EXE点灯
                                Case 11 : cISPComm.SendCmd = "TS,7,0,0,1620001A,2,0,"       'EXE消灯
                                Case 12 : cISPComm.SendCmd = "TS,7,0,0,1620001C,2,8000,"    'US1点灯
                                Case 13 : cISPComm.SendCmd = "TS,7,0,0,1620001C,2,0,"       'US1消灯
                                Case 14 : cISPComm.SendCmd = "TS,7,0,0,1620001E,2,8000,"    'US2点灯
                                Case 15 : cISPComm.SendCmd = "TS,7,0,0,1620001E,2,0,"       'US2消灯
                                Case 16 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,8,"       '8点灯
                                Case 17 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,0,"       '8消灯
                                Case 18 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,4,"       '4点灯
                                Case 19 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,0,"       '4消灯
                                Case 20 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,2,"       '2点灯
                                Case 21 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,0,"       '2消灯
                                Case 22 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,1,"       '1点灯
                                Case 23 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,0,"       '1消灯
                            End Select
                            Call cISPComm.FuncSendCmd()

                            'wait push button -------------------
                            If dfrm.DIALOG_RESULT = "NG" Then
                                mErrorComment = "LED 異常"
                                Call WriteLog(False, mErrorComment)
                                Throw New ApplicationException(mErrorComment)
                            End If
                            If dfrm.DIALOG_RESULT = "OK" Then
                                Exit Do
                            End If

                            DoEvents()
                            Call mSecWait(Timing)
                        Next i
                    Loop

                Catch ex As Exception
                    Throw New ApplicationException(mErrorComment)
                Finally
                    Call cISPComm.PortClose()
                    cISPComm = Nothing
                    dfrm.Close()
                    dfrm = Nothing
                End Try

                Call WriteLog(False, "LED検査 正常終了")
                ' ---------------------------------------------------------------------------------------
        End Select

        '全点灯確認
        Try
            'SP TestMode COM OPEN
            cISPComm = New cls101_ComRS232C
            cISPComm.RS_COMM = Insfrm.SerialPort1
            cISPComm.PORT = rs232cSPPort
            cISPComm.SETTING = "9600,E,8,1"
            cISPComm.TERMINAL_CHAR = vbCrLf
            cISPComm.PortSetup()
            For i = 0 To 8
                Select Case i
                    Case 0 : cISPComm.SendCmd = "TS,7,0,0,1660B21B,1,C0,"
                    Case 1 : cISPComm.SendCmd = "TS,7,0,0,1660B228,1,80,"
                    Case 2 : cISPComm.SendCmd = "TS,7,0,0,1660B229,1,80,"
                    Case 3 : cISPComm.SendCmd = "TS,7,0,0,1660B22A,1,80,"
                    Case 4 : cISPComm.SendCmd = "TS,7,0,0,16200018,2,8000,"
                    Case 5 : cISPComm.SendCmd = "TS,7,0,0,1620001A,2,8000,"
                    Case 6 : cISPComm.SendCmd = "TS,7,0,0,1620001C,2,8000,"
                    Case 7 : cISPComm.SendCmd = "TS,7,0,0,1620001E,2,8000,"
                    Case 8 : cISPComm.SendCmd = "TS,7,0,0,16200016,2,F,"
                End Select
                Call cISPComm.FuncSendCmd()
            Next

            Call mSecWait(500)
            If fCam = True Then
                Call cEI.GAZO_TEST()
                Call WriteLog(False, "全点灯")
                Call WriteLog(False, cEI.RESULT)
                'Call WriteLog(False, cEI.ResultData)
            End If

            'メッセージ表示
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\ALLLED.ppt"
            'dfrm.DISP_TITLENAME = "LED 検査"                                       'Rev0.11DEL
            dfrm.DISP_TITLENAME = "LED 검사"                                       'Rev0.11ADD
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "LED全点灯検査 異常 "
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
            cISPComm.SendCmd = "TS,28,0,0," : Call cISPComm.FuncSendCmd()
            Call cISPComm.PortClose()
            cISPComm = Nothing
            dfrm.Close()
            dfrm = Nothing
        End Try

        Call WriteLog(False, "LED検査 正常終了")

    End Sub

    Public Sub I23(ByVal iName As String)
        '==============================================================================================================
        ' Function : I23:USB検査-ﾌｧｰﾑ書き込み
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
        Call WriteLog(False, "USB検査－ﾌｧｰﾑ書込中・・・")

        Dim rcv As String
        Dim CopyWait As Integer
        Dim NewDriveName As String = ""

        'USB Connect -> New Drive
        Try
            If SubTest01(NewDriveName) = False Then
                mErrorComment = "USBﾓｰﾄﾞ　接続失敗"
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "ＵＳＢ接続完了")
        Call WriteLog(False, "ファームコピー中・・・")

        'Firmware Copy
        Try
            If SubTest02(NewDriveName) = False Then
                mErrorComment = "ﾌｧｰﾑｳｪｱのｺﾋﾟｰ失敗"
                Throw New ApplicationException(mErrorComment)
            End If

            CopyWait = 30
            Call WriteLog(False, "コピー後処理待ち " & CopyWait & "秒")

            Call WaitStart(30)

        Catch ex As Exception

            mErrorComment = "ﾌｧｰﾑｳｪｱのｺﾋﾟｰ失敗"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            Call WriteLog(False, "ＵＳＢ遮断中・・・")

            'USB DisConnect
            '送受信
            cISPComm = New cls101_ComRS232C
            cISPComm.RS_COMM = Insfrm.SerialPort1
            cISPComm.PORT = rs232cSPPort
            cISPComm.SETTING = "9600,N,8,1"
            cISPComm.TERMINAL_CHAR = vbCrLf

            Try
                cISPComm.SendCmd = "!"
                Call cISPComm.PortSetup()
                Call cISPComm.FuncSendCmd()

                Call WaitTimer(1)
                Call cISPComm.FuncRecvCmd()

                rcv = cISPComm.RECEIVE_DATA

                'If rcv <> "TS,81,0,0,0," Then
                '    mErrorComment = "!ｺﾏﾝﾄﾞ通信　異常"
                '    Throw New ApplicationException(mErrorComment)
                'End If
            Catch ex As Exception
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            Finally
                Call cISPComm.PortClose()
                cISPComm = Nothing
            End Try

        End Try

      
        'Try
        '    rcv = CommandComm("!", "Boot", 1)
        '    If rcv <> "!" Then
        '        mErrorComment = "!ｺﾏﾝﾄﾞ通信　異常"
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End Try

        'USB ﾒｯｾｰｼﾞ消去
        Call WaitTimer(1)
        'cWin = New cls001_Window_Control
        'cWin.WINDOW_TITLE = "デバイスの取り外しの警告"
        'Call cWin.WINDOW_QUIT()
        'cWin = Nothing

        Call WriteLog(False, "ＵＳＢ遮断完了")

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I24(ByVal iName As String)
        '==============================================================================================================
        ' Function : I24:SIO-制御線
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
        Call WriteLog(False, "SIO-制御線検査中・・・")

        Dim rcv As String

        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = vbCrLf
        Call cISPComm.PortSetup()

        Try
            'CTSﾈｹﾞｰﾄ ----------------------------
            WriteLog(False, "CTSﾈｹﾞｰﾄ中・・・約２秒待ち")
            cISPComm.SendCmd = "TS,6,0,0,00206084,4,"
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendCmd()
            Call cISPComm.FuncRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,6,0,0,0,4,4127,206084," Then
                mErrorComment = "CTSﾈｹﾞｰﾄ　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "CTSﾈｹﾞｰﾄ OK " & rcv)
            Call cISPComm.RecBufIni()
            'Call WaitStart(2)

            'RTSﾈｹﾞｰﾄ確認-------------------------
            Call WriteLog(False, "RTSﾈｹﾞｰﾄ確認中・・・約２秒待ち")
            cISPComm.SendCmd = "TS,6,0,0,00206094,4,"
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendCmd()
            Call cISPComm.FuncRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,6,0,0,0,4,2050,206094," Then
                mErrorComment = "RTSﾈｹﾞｰﾄ確認　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "RTSﾈｹﾞｰﾄ確認 OK " & rcv)
            Call cISPComm.RecBufIni()
            'Call WaitStart(2)

            'CTSｱｻｰﾄ -----------------------------
            Call WriteLog(False, "CTSｱｻｰﾄ中・・・約２秒待ち")
            cISPComm.SendCmd = "TS,7,0,0,00206084,4,00005127,"
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendCmd()
            Call cISPComm.FuncRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,7,0,0,0,4,206084," Then
                mErrorComment = "CTSｱｻｰﾄ　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "CTSｱｻｰﾄ OK " & rcv)
            Call cISPComm.RecBufIni()
            'Call WaitStart(2)

            'RTSｱｻｰﾄ -----------------------------
            Call WriteLog(False, "RTSｱｻｰﾄ中・・・約２秒待ち")
            cISPComm.SendCmd = "TS,6,0,0,00206094,4,"
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendCmd()
            Call cISPComm.FuncRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,6,0,0,0,4,7050,206094," Then
                'If rcv <> "TS,6,0,0,0,4,2050,206094," Then
                mErrorComment = "RTSｱｻｰﾄ　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "RTSｱｻｰﾄ OK " & rcv)
            Call cISPComm.RecBufIni()
            'Call WaitStart(2)

            'CTSﾈｹﾞｰﾄ ----------------------------
            Call WriteLog(False, "CTSﾈｹﾞｰﾄ中・・・約２秒待ち")
            cISPComm.SendCmd = "TS,7,0,0,00206084,4,00004127,"
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendCmd()
            Call cISPComm.FuncRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,7,0,0,0,4,206084," Then
                mErrorComment = "CTSﾈｹﾞｰﾄ　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "CTSﾈｹﾞｰﾄ OK " & rcv)
            Call cISPComm.RecBufIni()
            'Call WaitStart(2)

            'RTSﾈｹﾞｰﾄ確認-------------------------
            Call WriteLog(False, "RTSﾈｹﾞｰﾄ確認中・・・約２秒待ち")
            cISPComm.SendCmd = "TS,6,0,0,00206094,4,"
            Call WriteLog(False, "Send:" & cISPComm.SendCmd)
            Call cISPComm.FuncSendCmd()
            Call cISPComm.FuncRecvCmd()
            Call WriteLog(False, "Recv:" & cISPComm.RECEIVE_DATA)
            rcv = cISPComm.RECEIVE_DATA
            If rcv <> "TS,6,0,0,0,4,3050,206094," Then
                'If rcv <> "TS,6,0,0,0,4,2050,206094," Then
                mErrorComment = "RTSﾈｹﾞｰﾄ確認　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If

            Call WriteLog(False, "RTSﾈｹﾞｰﾄ確認 OK " & rcv)

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I25(ByVal iName As String)
        '==============================================================================================================
        ' Function : I25:SIO-ﾎﾞｰﾚｰﾄ
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
        Call WriteLog(False, "SIO-ﾎﾞｰﾚｰﾄ検査中・・・")

        Dim rcv As String
        Dim BRate As String = ""

        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.TIMEOUT = 10

        For i = 1 To 8
            Select Case i
                Case 1      '9600->1200bps
                    BRate = "1200" : cISPComm.SETTING = "9600,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,1,"
                Case 2      '1200->2400bps
                    BRate = "2400" : cISPComm.SETTING = "1200,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,2,"
                Case 3      '2400->4800bps
                    BRate = "4800" : cISPComm.SETTING = "2400,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,3,"
                Case 4      '4800->19200bps
                    BRate = "19200" : cISPComm.SETTING = "4800,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,5,"
                Case 5      '19200->38400bps
                    BRate = "38400" : cISPComm.SETTING = "19200,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,6,"
                Case 6      '38400->115200bps
                    BRate = "115200" : cISPComm.SETTING = "38400,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,7,"
                Case 7      '115200->57600bps
                    BRate = "57600" : cISPComm.SETTING = "115200,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,8,"
                Case 8      '57600->9600bps
                    BRate = "9600" : cISPComm.SETTING = "57600,E,8,1" : cISPComm.SendCmd = "TS,25,0,0,4,"
            End Select

            Call WriteLog(False, BRate & "bps変更中・・・約２秒待ち")
            Try
                Call cISPComm.PortSetup()
                Call WriteLog(False, "Send:" & cISPComm.SendCmd)
                cISPComm.TERMINAL_CHAR = vbCrLf
                Call cISPComm.FuncSendCmd()
                cISPComm.TERMINAL_CHAR = vbLf
                Call cISPComm.FuncRecvCmd()
                'Call cISPComm.FuncKYOUSEIRecvCmd()
                Call WriteLog(False, "Recv:" & Strings.Left(cISPComm.RECEIVE_DATA, 12))
                'rcv = CommandComm(cISPComm.SendCmd, "Test", 5)
                rcv = cISPComm.RECEIVE_DATA
                If Strings.Left(rcv, 12) <> "TS,25,0,0,0," Then
                    mErrorComment = BRate & "bps変更　異常 " & rcv
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, BRate & "bps変更 OK " & rcv)
                Call cISPComm.PortClose()

            Catch ex As Exception
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End Try

            'PC側設定変更
            Call WriteLog(False, "PC側 " & BRate & "bpsに変更")

        Next i

        cISPComm = Nothing
        Call WriteLog(False, "SIOﾎﾞｰﾚｰﾄ確認 OK ")

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I26(ByVal iName As String)
        '==============================================================================================================
        ' Function : I26:ﾗﾀﾞｰ検査
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
        'Dim FunctionName As String = "ﾗﾀﾞｰ検査"         'Rev0.11DEL
        Dim FunctionName As String = "Ladder검사"         'Rev0.11ADD

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, FunctionName & "中・・・")

        Dim IDevice As String
        Dim RecvDataLE As String = ""
        Dim Result As String
        Dim RunWait As Integer

        Call WriteLog(False, FunctionName & "-SDｱｸｾｽ/SET SW検査")

        'メッセージ表示 -------------------------------------------------
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\SDStart.ppt"
            dfrm.DISP_TITLENAME = FunctionName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "ラダー運転で中止が押されました。"
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

        Call WriteLog(False, "ラダー運転開始")

        RunWait = 5
        Call WriteLog(False, "ラダー運転 " & RunWait & "秒待ち")

        Call WaitStart(RunWait)

        Call WriteLog(False, FunctionName & "SPU-KP割込/DMA-RQ/SCB用信号検査中・・・")

        'Iレジスタ読み込み -------------------------------------
        IDevice = "I00001"

        Call WriteLog(False, "Iレジスタ読み込み")

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U2
            Call cLnk.F3PCLinkCommunication("02" & "WRD" & IDevice + ",01", RecvDataLE)

            Result = Mid(RecvDataLE, 3, 2)        'Strings.Left(RecvDataLE, 2)
            If Result <> "OK" Then
                mErrorComment = "Iレジスタ受信　異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If

            If Mid(RecvDataLE, 5, 4) <> "F00F" Then
                mErrorComment = FunctionName & "SPU-KP割込/DMA-RQ/SCB用信号検査　異常 " & Mid(RecvDataLE, 5, 4) & " (正常はF00F)"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, FunctionName & "SPU-KP割込/DMA-RQ/SCB用信号検査　正常 " & Mid(RecvDataLE, 5, 4))

        '運転停止 ----------------------------------------------
        Call WriteLog(False, "ラダー運転 停止")
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_SP
        cMcm.UnitNo = 0 : cMcm.SlotNo = 2
        cMcm.RemPort = 12290
        Call cMcm.SEQ_STOP()
        If cMcm.RecvText <> "0200200200020000" Then     '0000200200020000
            mErrorComment = FunctionName & "運転停止指令　異常 " & cMcm.RecvText
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "ラダー運転 停止 OK")
        
        Call mSecWait(500)

        'プログラム消去 ----------------------------------------------
        Call WriteLog(False, "ラダー 消去")
        Try
            Call cMcm.MEM_CLEAR()
            'If cMcm.RecvText <> "0200220100020000" Then
            '    mErrorComment = FunctionName & "運転停止指令　異常 " & cMcm.RecvText
            '    Throw New ApplicationException(mErrorComment)
            'End If

        Catch ex As Exception

        End Try

        Call WriteLog(False, "ラダー 消去 OK")
        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I27(ByVal iName As String)
        '==============================================================================================================
        ' Function : I27:SDｶｰﾄﾞ検査-ﾗｲﾄﾌﾟﾛﾃｸﾄ
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

        'Test
        ' ---------------------------------------------------------------------------------------
        Dim rcv As String

        ''WriteProtect無し確認 -------------------------
        'Call WriteLog(False, "WriteProtect無し確認・・・約２秒待ち")

        'Try
        '    rcv = CommandComm("TS,6,0,0,0021C224,4,", "Test", 5)
        '    '    If rcv <> "TS,6,0,0,0,4,A393FFF8,21C224," Then
        '    If Mid(rcv, 14, 4) <> "A393" Then
        '        mErrorComment = "WriteProtect無し確認　異常 " & rcv
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End Try

        'Call WriteLog(False, "WriteProtect無し確認 OK " & rcv)

        ''メッセージ表示 -------------------------------------------------
        'Try
        '    dfrm = New frm500_MessagePopUp
        '    dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\SDProtect.ppt"
        '    dfrm.DISP_TITLENAME = iName
        '    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        '    Call dfrm.frmShow()

        '    Do
        '        'wait push button -------------------
        '        If dfrm.DIALOG_RESULT = "NG" Then
        '            mErrorComment = "WriteProtect確認で中止が押されました。"
        '            Throw New ApplicationException(mErrorComment)
        '        End If

        '        If dfrm.DIALOG_RESULT = "OK" Then
        '            Exit Do
        '        End If

        '        DoEvents()

        '        Call mSecWait(200)

        '    Loop

        'Catch ex As Exception
        '    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
        '    mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)

        'Finally
        '    dfrm.Close()
        '    dfrm = Nothing
        'End Try

        'WriteProtect有り確認 -------------------------
        Call WriteLog(False, "WriteProtect有り確認・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,6,0,0,0021C224,4,", "Test", 5)
            '    If rcv <> "TS,6,0,0,0,4,A3D3FFF8,21C224," Then
            If Mid(rcv, 14, 4) <> "A3D3" Then
                mErrorComment = "WriteProtect有り確認　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "WriteProtect有り確認 OK " & rcv)

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
        Dim TestStd_SDPWR_H As Single = 3.38
        Dim TestStd_SDPWR_L As Single = 3
        Dim TestStd_SDDWN_H As Single = 0.2
        Dim TestStd_SDDWN_L As Single = 0
        Dim TestStd_VCCSDI_H As Single = 3.38
        Dim TestStd_VCCSDI_L As Single = 3.05

        Dim cMeas As New cls140_VoltageMeasure

        ''メッセージ表示 -------------------------------------------------
        'Try
        '    dfrm = New frm500_MessagePopUp
        '    dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\SDChange.ppt"
        '    dfrm.DISP_TITLENAME = iName
        '    dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

        '    Call dfrm.frmShow()

        '    Do
        '        'wait push button -------------------
        '        If dfrm.DIALOG_RESULT = "NG" Then
        '            mErrorComment = "中止が押されました。"
        '            Throw New ApplicationException(mErrorComment)
        '        End If

        '        If dfrm.DIALOG_RESULT = "OK" Then
        '            Exit Do
        '        End If

        '        DoEvents()

        '        Call mSecWait(200)

        '    Loop

        'Catch ex As Exception
        '    mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
        '    mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)

        'Finally
        '    dfrm.Close()
        '    dfrm = Nothing
        'End Try

        'VCCSDI -------------------------------------
        SigName = "VCCSDI"
        ADDevice = "0001"
        StdMax = TestStd_VCCSDI_H : StdMin = TestStd_VCCSDI_L

        Call WriteLog(False, SigName & "検査中・・・")

        'cLnk = New cls080_F3LinkToLE
        'cLnk.IPAdd = IP_U1
        'Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
        'RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
        'Result = Strings.Left(RecvDataLE, 2)
        'If Result <> "OK" Then
        '    mErrorComment = SigName & "測定返信 異常 " & RecvDataLE
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'End If
        'sValue = Mid(RecvDataLE, 3, 4)
        'dValue = CDec("&h" & sValue) / (20000 / 10)
        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)

        'SD電源 -------------------------------------
        SigName = "SDPWR"
        ADDevice = "0007"
        StdMax = TestStd_SDPWR_H : StdMin = TestStd_SDPWR_L

        Call WriteLog(False, SigName & "検査中・・・")

        'cLnk = New cls080_F3LinkToLE
        'cLnk.IPAdd = IP_U1
        'Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)

        'Result = Mid(RecvDataLE, 3, 2)        'Strings.Left(RecvDataLE, 2)
        'If Result <> "OK" Then
        '    mErrorComment = SigName & "測定返信 異常 " & RecvDataLE
        '    Throw New ApplicationException(mErrorComment)
        'End If
        'sValue = Mid(RecvDataLE, 5, 4)
        'dValue = CDec("&h" & sValue) / (20000 / 10)
        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        Call WriteValueData("VCCSDI", dValue.ToString)

        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)

        'ResultData.SDPower = dValue

        '電源ショート -------------------------------------
        YDDevice = "903"
        cLnk = New cls080_F3LinkToLE
        cLnk.IPAdd = IP_U1
        Call cLnk.F3PCLinkCommunication("01" & "BWRY" + YDDevice + ",001,1", RecvDataLE)

        If Mid(RecvDataLE, 3, 2) <> "OK" Then
            mErrorComment = "SD電源ｼｮｰﾄ 異常 "
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "SD電源ｼｮｰﾄ 正常")

        Call WaitTimer(1)

        YDDevice = "903"
        Call cLnk.F3PCLinkCommunication("01" & "BWRY" + YDDevice + ",001,0", RecvDataLE)

        If Mid(RecvDataLE, 3, 2) <> "OK" Then
            mErrorComment = "SD電源ｼｮｰﾄ解除 異常 "
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "SD電源ｼｮｰﾄ解除 正常")

        'SD電源 -------------------------------------
        SigName = "SDDWN"
        ADDevice = "0007"
        StdMax = TestStd_SDDWN_H : StdMin = TestStd_SDDWN_L

        Call WriteLog(False, SigName & "検査中・・・")

        'Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)

        'Result = Mid(RecvDataLE, 3, 2)        'Strings.Left(RecvDataLE, 2)
        'If Result <> "OK" Then
        '    mErrorComment = SigName & "測定返信 異常 " & RecvDataLE
        '    Throw New ApplicationException(mErrorComment)
        'End If
        'sValue = Mid(RecvDataLE, 5, 4)
        'dValue = CDec("&h" & sValue) / (20000 / 10)
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)

        'メッセージ表示 -------------------------------------------------
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\SDChange2.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "ｼｮｰﾄ後のSD認識　異常"
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
        ' Function : I29:Ethernet通信検査
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

        'Call WriteLog(False, "開始待ち・・・5秒")
        'Call WaitTimer(5)


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

        'Call WriteLog(False, "速度変更　反映待機中・・・20秒")     'v0.02ｺﾒﾝﾄ
        'Call WaitTimer(20)
        Call WriteLog(False, "速度変更　反映待機中・・・")       'v0.02追加
        Call WaitStart(7)

        'FTP→Mｺﾏﾝﾄﾞ通信に変更　v0.02
        Call WriteLog(False, "通信速度10M Fullでの通信確認中・・・")
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_SP : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 0 : cMcm.TimeOut = 5
        Dim i As Integer = 0
        Do
            Call cMcm.RTC_SET()     'RTC書き込み(現時刻)
            If cMcm.RecvText = "0000021200020000" Then Exit Do
            If i >= 1 Then
                mErrorComment = iName & "RTC設定　返信　異常 " & cMcm.RecvText
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            i = i + 1
            Call WriteLog(False, "RTC設定 ﾘﾄﾗｲ待ち　5秒")
            Call WaitStart(5)
            DoEvents()
        Loop
        cMcm = Nothing

        Call WriteLog(False, "Mｺﾏﾝﾄﾞ通信10M OK")
        
        Call WriteLog(False, "通信速度10M Fullでの通信　OK")
        Call WriteLog(False, "")

        Call WaitStart(1)

        '速度100M Fullに変更 -----------------------------------------------------
        Call WriteLog(False, "通信速度100M Fullに変更中・・・")

        RegStr = "4" '100M Full

        cPCSP = New cls060_PCNetSPEED

        Ret1 = cPCSP.VBSetRegistryString(RegStr, sKEY)
        Ret = cPCSP.DisEnableNetAdapter

        'Call WriteLog(False, "速度変更　反映待機中・・・20秒")     'v0.02ｺﾒﾝﾄ
        'Call WaitTimer(20)
        Call WriteLog(False, "速度変更　反映待機中・・・")       'v0.02追加
        'Call WaitStart(5)             Rev0.12DEL

        '-----< Rev0.12ADD >-----
        For i = 0 To 10
            If My.Computer.Network.Ping("192.168.0.2") = False Then
                Call WaitStart(1)
            Else
                Exit For
            End If
            If i = 10 Then
                mErrorComment = "PING 通信異常"
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
        Next i
        '------------------------

        'FTP
        Call WriteLog(False, "通信速度100M Fullでの通信確認中・・・")
        Try
            If ReWriteSD() = False Then
                mErrorComment = "Ethernet通信検査(FTP)SD書換え 異常"
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "Ethernet通信検査(FTP)SD書換え 異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "通信速度100M Fullでの通信（SD書換え）　OK")

        '正常終了 ----------------------------------------------------------------
        Call WriteLog(False, "Ethernet通信検査 完了")

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I30(ByVal iName As String)
        '==============================================================================================================
        ' Function : I30:SW検査
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
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\SW.ppt"
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

    Public Sub I31(ByVal iName As String)
        '==============================================================================================================
        ' Function : I31:電源電圧検査
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

        Dim TestStd_VCC_H As Single = 5.15
        Dim TestStd_VCC_L As Single = 4.85
        Dim TestStd_VCC3_H As Single = 3.29
        Dim TestStd_VCC3_L As Single = 3.14
        Dim TestStd_V1P9_H As Single = 1.99
        Dim TestStd_V1P9_L As Single = 1.8
        Dim TestStd_V2P5_H As Single = 2.62
        Dim TestStd_V2P5_L As Single = 2.38

        Dim cMeas As New cls140_VoltageMeasure

        Try
            For i = 1 To 4
                Select Case i
                    Case 1
                        SigName = "VCC" : ADDevice = "0006" : StdMax = TestStd_VCC_H : StdMin = TestStd_VCC_L
                    Case 2
                        SigName = "VCC3" : ADDevice = "0002" : StdMax = TestStd_VCC3_H : StdMin = TestStd_VCC3_L
                    Case 3
                        SigName = "V1P9" : ADDevice = "0004" : StdMax = TestStd_V1P9_H : StdMin = TestStd_V1P9_L
                    Case 4
                        SigName = "V2P5" : ADDevice = "0005" : StdMax = TestStd_V2P5_H : StdMin = TestStd_V2P5_L
                End Select

                Call WriteLog(False, SigName & "検査中・・・")
                'v0.07　DMM測定
                dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))
                Call WriteValueData(SigName, dValue.ToString)

                If dValue > StdMax Or dValue < StdMin Then
                    mErrorComment = SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
                    Call WriteLog(False, mErrorComment)
                    Throw New ApplicationException(mErrorComment)
                End If

                Call WriteLog(False, SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)

            Next i

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMeas = Nothing
        End Try

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I32(ByVal iName As String)
        '==============================================================================================================
        ' Function : I32:ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟ検査
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
        Call WriteLog(False, "ﾊﾞｯﾃﾘﾊﾞｯｸｱｯﾌﾟ検査中・・・")

        Dim ADDevice As String
        Dim RecvDataLE As String = ""
        'Dim Result As String
        'Dim sValue As String
        Dim dValue As Single
        Dim StdMax As Single
        Dim StdMin As Single
        Dim SigName As String
        Dim rcv As String
        Dim FuncName As String
        Dim PwrONWait As Integer
        Dim PwrOFFWait As Integer
        Dim PCTime As String
        Dim RTCTime As String
        Dim i As Integer

        Dim TestStd_VC3SLP_ON_H As Single = 3.29
        Dim TestStd_VC3SLP_ON_L As Single = 2.94
        Dim TestStd_VC3SLP_OFF_H As Single = 3.15   'v0.08 3->3.15
        Dim TestStd_VC3SLP_OFF_L As Single = 2.3    'v0.08 2.4->2.3

        Dim ResultData_DCVC3SLP_ON As Double
        Dim ResultData_DCVC3SLP_OFF As Double
        Dim ResultData_DCD_VC3SLP As Double
        Dim ResultData_M1D_VC3SLP As Double
        Dim ResultData_VC3SLP_ON As Double
        Dim ResultData_M1VC3SLP_OFF As Double
        Dim ResultData_M0VC3SLP_OFF As Double
        Dim ResultData_M0D_VC3SLP As Double

        Dim cMeas As New cls140_VoltageMeasure

        PwrONWait = 5
        PwrOFFWait = 3

        'ﾒﾓﾘ0 ****************************************************************************
        Call WriteLog(False, "ﾒﾓﾘ 0 時検査 *************")
        'VC3SLP -------------------------------------
        SigName = "VC3SLP"
        ADDevice = "0003"
        StdMax = TestStd_VC3SLP_ON_H : StdMin = TestStd_VC3SLP_ON_L

        Call WriteLog(False, "電源ON時" & SigName & "測定中・・・")

        'cLnk = New cls080_F3LinkToLE
        'Try
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

        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = "電源ON時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)
        Call WriteValueData(SigName & "_ON", dValue.ToString)

        ResultData_VC3SLP_ON = dValue

        'RTC設定 -------------------------------------
        Call WriteLog(False, "RTC設定")

        Call WriteLog(False, "設定開始待ち・・・")
        Call WaitStart(3)

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_SP : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 0 : cMcm.TimeOut = 5
        i = 0
        Do
            Try
                Call cMcm.RTC_SET()     'RTC書き込み(現時刻)
            Catch ex As Exception
                mErrorComment = "RTC書き込み失敗。Ethernet通信異常。"
                Call WriteLog(False, mErrorComment)
                cMcm = Nothing
                Throw New ApplicationException(mErrorComment)
            End Try
            RecvDataLE = cMcm.RecvText
            If RecvDataLE = "0000021200020000" Then Exit Do

            If i >= 1 Then
                mErrorComment = iName & "RTC設定　返信　異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If
            i = i + 1
            Call WriteLog(False, "RTC設定 ﾘﾄﾗｲ待ち　5秒")
            Call WaitStart(5)
            DoEvents()
        Loop

        cMcm = Nothing

        Call WriteLog(False, "RTC設定 OK" & RecvDataLE)

        '--------------------
        FuncName = "ﾃｽﾄﾓｰﾄﾞ移行"
        Call WriteLog(False, FuncName & "・・・")

        If TestModeStart() = False Then
            mErrorComment = FuncName & "　異常"
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " " & PwrOFFWait & "秒待ち")
        Call WaitStart(PwrOFFWait)
        Call WriteLog(False, "テストモード移行　完了")

        'ﾒﾓﾘｸﾘｱ --------------------------------------
        Call WriteLog(False, "ﾒﾓﾘｸﾘｱ0")
        '--------------------
        FuncName = "共有ﾒﾓﾘ0書込"
        Call WriteLog(False, FuncName & "・・・約２秒待ち")

        rcv = CommandComm("TS,21,0,0,16800000,1687FFFF,4,0,", "Test", 5)
        If rcv <> "TS,21,0,0,0,1687FFFC," Then
            mErrorComment = FuncName & " 異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)
        '--------------------
        FuncName = "DVBﾒﾓﾘ指定"
        Call WriteLog(False, FuncName & "・・・約２秒待ち")

        rcv = CommandComm("TS,7,0,0,16F00008,2,7,", "Test", 5)
        If rcv <> "TS,7,0,0,0,2,16F00008," Then
            mErrorComment = FuncName & " 異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)
        '--------------------
        FuncName = "DVBﾒﾓﾘ0書込"
        Call WriteLog(False, FuncName & "・・・約３秒待ち")

        rcv = CommandComm("TS,21,0,0,16C00000,16CFFFFF,4,0,", "Test", 5)
        If rcv <> "TS,21,0,0,0,16CFFFFC," Then
            mErrorComment = FuncName & " 異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '--------------------
        FuncName = "ﾃｽﾄﾓｰﾄﾞ終了"
        Call WriteLog(False, FuncName & "・・・")

        rcv = CommandComm("TS,2,0,0,", "Test", 5)
        If rcv <> "TS,2,0,0,0," Then
            mErrorComment = "テストモード終了　異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "TestMode終了　2秒待ち")
        Call WaitStart(2)
        Call WriteLog(False, "TestMode終了　完了")

        '電源遮断 -------------------------------------
        Call WriteLog(False, "電源遮断")

        Call I91(iName)

        Call WriteLog(False, "電源遮断 " & PwrOFFWait & "秒待ち")
        Call WaitStart(PwrOFFWait)

        'VC3SLP -------------------------------------
        SigName = "VC3SLP"
        ADDevice = "0003"
        StdMax = TestStd_VC3SLP_OFF_H : StdMin = TestStd_VC3SLP_OFF_L

        Call WriteLog(False, "電源OFF時" & SigName & "測定中・・・")

        'cLnk = New cls080_F3LinkToLE
        'Try
        '    cLnk.IPAdd = IP_U1
        '    Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice & ",01", RecvDataLE)
        '    RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
        '    Result = Strings.Left(RecvDataLE, 2)
        '    If Result <> "OK" Then
        '        mErrorComment = "電源OFF時" & SigName & "測定返信 異常 " & RecvDataLE
        '        Call WriteLog(False, mErrorComment)
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    sValue = Mid(RecvDataLE, 3, 4)
        '    dValue = CDec("&h" & sValue) / (20000 / 10)
        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    cLnk = Nothing
        'End Try
        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        'If dValue < StdMin Then dValue += CSng(0.07)
        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = "電源OFF時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "電源OFF時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)
        Call WriteValueData(SigName & "_OFF", dValue.ToString)
        ResultData_M0VC3SLP_OFF = dValue

        'VC3SLP差分(ﾒﾓﾘ0) -------------------------------------
        ResultData_M0D_VC3SLP = ResultData_VC3SLP_ON - ResultData_M0VC3SLP_OFF
        Call WriteLog(False, "ON時電圧-OFF時電圧=" & CStr(ResultData_M0D_VC3SLP))
        Call WriteLog(False, "")

        'ﾒﾓﾘ1 ****************************************************************************
        Call WriteLog(False, "ﾒﾓﾘ 1 時検査 *************")

        '電源投入 -------------------------------------
        Call I90(iName)

        Call WriteLog(False, "電源投入 " & PwrONWait & "秒待ち")

        Call WaitStart(PwrONWait)

        '--------------------
        FuncName = "ﾃｽﾄﾓｰﾄﾞ移行"
        Call WriteLog(False, FuncName & "・・・")

        If TestModeStart() = False Then
            mErrorComment = FuncName & "　異常" : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " " & PwrOFFWait & "秒待ち")
        Call WaitStart(PwrOFFWait)
        Call WriteLog(False, "テストモード移行　完了")

        'ﾒﾓﾘｸﾘｱ --------------------------------------
        Call WriteLog(False, "ﾒﾓﾘｸﾘｱ1")

        '--------------------
        FuncName = "共有ﾒﾓﾘ1書込"
        Call WriteLog(False, FuncName & "・・・約２秒待ち")

        rcv = CommandComm("TS,21,0,0,16800000,1687FFFF,4,FFFFFFFF,", "Test", 5)
        If rcv <> "TS,21,0,0,0,1687FFFC," Then
            mErrorComment = FuncName & " 異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '--------------------
        FuncName = "DVBﾒﾓﾘ指定"
        Call WriteLog(False, FuncName & "・・・約２秒待ち")

        rcv = CommandComm("TS,7,0,0,16F00008,2,7,", "Test", 5)
        If rcv <> "TS,7,0,0,0,2,16F00008," Then
            mErrorComment = FuncName & " 異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If


        Call WriteLog(False, FuncName & " OK " & rcv)

        '--------------------
        FuncName = "DVBﾒﾓﾘ1書込"
        Call WriteLog(False, FuncName & "・・・約３秒待ち")

        rcv = CommandComm("TS,21,0,0,16C00000,16CFFFFF,4,FFFFFFFF,", "Test", 5)
        If rcv <> "TS,21,0,0,0,16CFFFFC," Then
            mErrorComment = FuncName & " 異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '--------------------
        FuncName = "ﾃｽﾄﾓｰﾄﾞ終了"
        Call WriteLog(False, FuncName & "・・・")

        rcv = CommandComm("TS,2,0,0,", "Test", 5)
        If rcv <> "TS,2,0,0,0," Then
            mErrorComment = "テストモード終了　異常 " & rcv : Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "TestMode終了　2秒待ち")

        Call WaitStart(2)
        Call WriteLog(False, "TestMode終了　完了")

        '電源遮断 -------------------------------------
        Call WriteLog(False, "電源遮断")

        Call I91(iName)

        Call WriteLog(False, "電源遮断 " & PwrOFFWait & "秒待ち")
        Call WaitStart(PwrOFFWait)

        'VC3SLP -------------------------------------
        SigName = "VC3SLP"
        ADDevice = "0003"
        StdMax = TestStd_VC3SLP_OFF_H : StdMin = TestStd_VC3SLP_OFF_L

        Call WriteLog(False, "電源OFF時" & SigName & "測定中・・・")

        'cLnk = New cls080_F3LinkToLE
        'Try
        '    cLnk.IPAdd = IP_U1
        '    Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice + ",01", RecvDataLE)
        '    RecvDataLE = Strings.Mid(RecvDataLE, 3, RecvDataLE.Length - 2)
        '    Result = Strings.Left(RecvDataLE, 2)
        '    If Result <> "OK" Then
        '        mErrorComment = "電源OFF時" & SigName & "測定返信 異常 " & RecvDataLE
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    sValue = Mid(RecvDataLE, 3, 4)
        '    dValue = CDec("&h" & sValue) / (20000 / 10)
        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    cLnk = Nothing
        'End Try
        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        'If dValue < StdMin Then dValue += CSng(0.07)
        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = "電源OFF時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "電源OFF時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)


        ResultData_M1VC3SLP_OFF = dValue

        'VC3SLP差分(ﾒﾓﾘ1) -------------------------------------
        ResultData_M1D_VC3SLP = ResultData_VC3SLP_ON - ResultData_M1VC3SLP_OFF
        Call WriteLog(False, "ON時電圧-OFF時電圧=" & CStr(ResultData_M1D_VC3SLP))
        Call WriteLog(False, "")


        '通常 ****************************************************************************
        Call WriteLog(False, "通常時検査 *************")

        '電源投入 -------------------------------------
        Call WriteLog(False, "電源投入")

        Call I90(iName)

        Call WriteLog(False, "電源投入 " & PwrONWait & "秒待ち")
        Call WaitStart(PwrONWait)

        '--------------------
        FuncName = "ﾊﾞｯｸｱｯﾌﾟﾃﾞｰﾀ書込準備"
        Call WriteLog(False, FuncName & "開始待ち 10秒")
        Call WaitStart(10)

        cLnk = New cls080_F3LinkToLE
        Try
            cLnk.IPAdd = IP_SP
            Call cLnk.F3PCLinkCommunication("02MCM100100080002000116805FC8", RecvDataLE)
            If Mid(RecvDataLE, 3, 18) <> "OK100100040000A55A" Then
                mErrorComment = FuncName & " 異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, FuncName & " OK " & RecvDataLE)

        '--------------------
        FuncName = "ﾊﾞｯｸｱｯﾌﾟﾃﾞｰﾀ書込"
        Call WriteLog(False, FuncName)

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_SP
            Call cLnk.F3PCLinkCommunication("02MCM1002000A0002000116805FB8A55A", RecvDataLE)
            If Mid(RecvDataLE, 3, 14) <> "OK100200020000" Then
                mErrorComment = FuncName & " 異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, FuncName & " OK " & RecvDataLE)

        '--------------------
        FuncName = "ﾊﾞｯｸｱｯﾌﾟﾃﾞﾊﾞｲｽ書込"
        Call WriteLog(False, FuncName)

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_SP
            Call cLnk.F3PCLinkCommunication("02" & "WWRD00001,01,55AA", RecvDataLE)
            If Mid(RecvDataLE, 1, 4) <> "12OK" Then
                mErrorComment = FuncName & " 異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try
        Call WriteLog(False, FuncName & " OK " & RecvDataLE)

        '電源遮断 -------------------------------------
        Call WriteLog(False, "電源遮断")

        Call I91(iName)

        Call WriteLog(False, "電源遮断 " & PwrOFFWait & "秒待ち")

        Call WaitStart(PwrOFFWait)

        'VC3SLP -------------------------------------
        SigName = "VC3SLP"
        ADDevice = "0003"
        StdMax = TestStd_VC3SLP_OFF_H : StdMin = TestStd_VC3SLP_OFF_L

        Call WriteLog(False, "電源OFF時" & SigName & "測定中・・・")

        'Try
        '    cLnk = New cls080_F3LinkToLE
        '    cLnk.IPAdd = IP_U1
        '    Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice + ",01", RecvDataLE)

        '    Result = Mid(RecvDataLE, 3, 2)
        '    If Result <> "OK" Then
        '        mErrorComment = "電源OFF時" & SigName & "測定返信 異常 " & RecvDataLE
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    sValue = Mid(RecvDataLE, 5, 4)
        '    dValue = CDec("&h" & sValue) / (20000 / 10)
        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    cLnk = Nothing
        'End Try
        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        'If dValue < StdMin Then dValue += CSng(0.07)
        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = "電源OFF時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "電源OFF時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)
        ResultData_DCVC3SLP_OFF = dValue

        'OFF時　Wait -------------------------------------
        Call WriteLog(False, "")
        Call WriteLog(False, "放電待ち 1分30秒待ち")
        Call WriteLog(False, "")

        'Call WaitTimer(90)
        Call WaitStart(90)

        '電源投入 -------------------------------------
        Call WriteLog(False, "電源投入")

        Call I90(iName)

        Call WriteLog(False, "電源投入 " & PwrONWait & "秒待ち")
        Call WaitStart(PwrONWait)

        'VC3SLP -------------------------------------
        SigName = "VC3SLP"
        ADDevice = "0003"
        StdMax = TestStd_VC3SLP_ON_H : StdMin = TestStd_VC3SLP_ON_L

        Call WriteLog(False, "電源ON時" & SigName & "測定中・・・")

        'Try
        '    cLnk = New cls080_F3LinkToLE
        '    cLnk.IPAdd = IP_U1
        '    Call cLnk.F3PCLinkCommunication("01" & "SWR" & "0" & "07" & "," & ADDevice + ",01", RecvDataLE)

        '    Result = Strings.Mid(RecvDataLE, 3, 2)
        '    If Result <> "OK" Then
        '        mErrorComment = "電源ON時" & SigName & "測定返信 異常 " & RecvDataLE
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    sValue = Mid(RecvDataLE, 5, 4)
        '    dValue = CDec("&h" & sValue) / (20000 / 10)
        'Catch ex As Exception
        '    Call WriteLog(False, mErrorComment)
        '    Throw New ApplicationException(mErrorComment)
        'Finally
        '    cLnk = Nothing
        'End Try
        'v0.07　DMM測定
        dValue = cMeas.VoltMeasure(Strings.Right(ADDevice, 1))

        If dValue > StdMax Or dValue < StdMin Then
            mErrorComment = "電源ON時" & SigName & "測定 異常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "電源ON時" & SigName & "測定 正常 測定値=" & dValue & "　規格：" & StdMin & "～" & StdMax)


        ResultData_DCVC3SLP_ON = dValue

        'VC3SLP差分() -------------------------------------
        ResultData_DCD_VC3SLP = ResultData_DCVC3SLP_ON - ResultData_DCVC3SLP_OFF
        Call WriteLog(False, "ON時電圧-OFF時電圧=" & CStr(ResultData_DCD_VC3SLP))
        Call WriteLog(False, "")


        'RTC確認 -------------------------------------
        Call WriteLog(False, "RTC確認")


        Call WriteLog(False, "読込開始待ち・・・")

        Call WaitStart(3)

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_SP : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 0 : cMcm.TimeOut = 5
        Dim j As Integer = 0
        Do
            Call cMcm.RTC_READ()        '"RTC読込み"
            RecvDataLE = cMcm.RecvText
            RTCTime = ""
            For i = 17 To 42 Step 2
                RTCTime = RTCTime & Chr(CInt("&H" & Mid(RecvDataLE, i, 2)))
            Next i

            PCTime = Format(DateTime.Now, "yy/MM/dd") & Format(DateTime.Now, "HH:mm")

            If RTCTime = PCTime Then Exit Do

            Call WriteLog(False, "RTC確認1回目失敗 SP:" & RTCTime & " PC:" & PCTime)

            If j >= 1 Then
                mErrorComment = "RTC確認2回目失敗 SP:" & RTCTime & " PC:" & PCTime
                Throw New ApplicationException(mErrorComment)
            End If

            j = j + 1
            Call WriteLog(False, "2回目読込開始待ち・・・10秒")

            Call WaitStart(10)

            DoEvents()

        Loop

        'ResultData.RTC_Read = RecvDataLE
        Call WriteLog(False, "RTC読込 OK : " & RecvDataLE & " PCTime:" & PCTime)

        '--------------------
        FuncName = "ﾊﾞｯｸｱｯﾌﾟﾃﾞｰﾀ確認"
        Call WriteLog(False, FuncName)

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_SP
            Call cLnk.F3PCLinkCommunication("02MCM100100080002000116805FB8", RecvDataLE)
            If Mid(RecvDataLE, 3, 18) <> "OK100100040000A55A" Then
                mErrorComment = FuncName & " 異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, FuncName & " OK " & RecvDataLE)

        '--------------------
        FuncName = "ﾊﾞｯｸｱｯﾌﾟﾃﾞﾊﾞｲｽ確認"
        Call WriteLog(False, FuncName)

        Try
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_SP
            Call cLnk.F3PCLinkCommunication("02" & "WRDD00001,01", RecvDataLE)
            If Mid(RecvDataLE, 3, 6) <> "OK55AA" Then
                mErrorComment = FuncName & " 異常 " & RecvDataLE
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cLnk = Nothing
        End Try

        Call WriteLog(False, FuncName & " OK " & RecvDataLE)

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I33(ByVal iName As String)
        '==============================================================================================================
        ' Function : I33:EthernetLED検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 10
        Dim rcv As String
        Dim rCmd As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "EthernetLED検査中・・・")

        'Call WriteLog(False, "TestMode終了 10秒待ち")
        'Call WaitStart(10)

        'Dim rcv As String

        'メッセージ表示
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\LAN LED1.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "EthernetLED検査 異常 "
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Exit Do
                End If

                DoEvents()

                Call mSecWait(100)

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

        'TestMode終了
        rcv = CommandComm("TS,2,0,0,", "Test", 5)
        If rcv <> "TS,2,0,0,0," Then
            mErrorComment = "テストモード終了　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "TestMode終了 15秒待ち")
        Call WaitStart(15)

        'メッセージ表示
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\LAN LED2.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_SP

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "EthernetLED検査 異常 "
                    Throw New ApplicationException(mErrorComment)
                End If

                If dfrm.DIALOG_RESULT = "OK" Then
                    Call WriteLog(False, "EthernetLED検査 正常")
                    Exit Do
                End If

                'If FTPConnect("/RAMDISK") = False Then          'Ethernet通信方法検討する。。。FTPすぐに接続できない。。。
                '    mErrorComment = "EthernetLED検査(FTP) 異常"
                '    Throw New ApplicationException(mErrorComment)
                'End If

                Try
                    Call cLnk.F3PCLinkCommunication("02WRDZ00121,03", rCmd)
                Catch ex As Exception
                    mErrorComment = "EthernetLED検査(PCﾘﾝｸ) 異常"
                    Throw New ApplicationException(mErrorComment)
                End Try

                DoEvents()

                Call mSecWait(100)

            Loop

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            dfrm.Close()
            dfrm = Nothing
            cLnk = Nothing
        End Try

        ' ---------------------------------------------------------------------------------------

    End Sub

    Public Sub I34(ByVal iName As String)
        '==============================================================================================================
        ' Function : I34:ｺﾏﾝﾄﾞ割込み検査
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
        Call WriteLog(False, "検査中・・・")


        Dim FuncName As String = ""
        Dim rcv As String
        Dim i As Integer
        Dim SendCom As String = ""
        Dim RecvCom As String = ""

        For i = 0 To 5
            Select Case i
                Case 0 : FuncName = "KP側ｺﾏﾝﾄﾞ割込みｲﾈｰﾌﾞﾙ" : SendCom = "TS,7,0,0,1660B21E,1,8," : RecvCom = "TS,7,0,0,0,1,1660B21E,"
                Case 1 : FuncName = "CPU側割込みﾏｽｸ" : SendCom = "TS,7,0,0,0021C030,4,0," : RecvCom = "TS,7,0,0,0,4,21C030,"
                Case 2 : FuncName = "CPU側割込み信号ｽﾃｰﾀｽ読込み" : SendCom = "TS,6,0,0,0021C024,4," : RecvCom = ""
                Case 3 : FuncName = "CPUｺﾏﾝﾄﾞ割込み" : SendCom = "TS,7,0,0,1660B223,1,80," : RecvCom = "TS,7,0,0,0,1,1660B223,"
                Case 4 : FuncName = "CPU側割込み信号ｽﾃｰﾀｽ読込み" : SendCom = "TS,6,0,0,0021C024,4," : RecvCom = ""
                Case 5 : FuncName = "CPUｺﾏﾝﾄﾞ割込みｸﾘｱ" : SendCom = "TS,7,0,0,1660B223,1,0," : RecvCom = "TS,7,0,0,0,1,1660B223,"
            End Select
            Call WriteLog(False, FuncName & "中・・・約２秒待ち")

            Try
                rcv = CommandComm(SendCom, "Test", 5)
                If RecvCom = "" Then
                    If i = 2 Then
                        If Mid(rcv, 20, 1) <> "F" Then
                            mErrorComment = FuncName & "　異常 " & rcv
                            Throw New ApplicationException(mErrorComment)
                        End If
                    End If
                    If i = 4 Then
                        If Mid(rcv, 20, 1) <> "B" Then
                            mErrorComment = FuncName & "　異常 " & rcv
                            Throw New ApplicationException(mErrorComment)
                        End If
                    End If
                Else
                    If rcv <> RecvCom Then
                        mErrorComment = FuncName & "　異常 " & rcv
                        Throw New ApplicationException(mErrorComment)
                    End If
                End If

            Catch ex As Exception
                Throw New ApplicationException(mErrorComment)
            End Try

            Call WriteLog(False, FuncName & " OK " & rcv)

        Next i

        ' ---------------------------------------------------------------------------------------


    End Sub

    Public Sub I35(ByVal iName As String)
        '==============================================================================================================
        ' Function : I35:RTC設定
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
        Dim PCTime As String
        Dim RTCTime As String
        Dim RecvDataLE As String = ""
        Dim i As Integer


        'RTC設定 -------------------------------------
        'Call WriteLog(False, "RTC設定")

        'Call WriteLog(False, "設定開始待ち・・・5秒")
        'Call WaitTimer(5)

        'cMcm = New cls120_SP_M_Command
        'cMcm.IPAdd = IP_SP : cMcm.RemPort = 12290 : cMcm.UnitNo = 0 : cMcm.SlotNo = 0 : cMcm.TimeOut = 5
        'i = 0
        'Do
        '    Call cMcm.RTC_SET()     'RTC書き込み(現時刻)
        '    RecvDataLE = cMcm.RecvText
        '    If RecvDataLE = "0000021200020000" Then Exit Do

        '    If i >= 1 Then
        '        mErrorComment = iName & "RTC設定　返信　異常 " & RecvDataLE
        '        Throw New ApplicationException(mErrorComment)
        '    End If
        '    i = i + 1
        '    Call WriteLog(False, "RTC設定 ﾘﾄﾗｲ待ち　5秒")
        '    Call WaitTimer(5)
        '    DoEvents()
        'Loop

        'cMcm = Nothing

        'Call WriteLog(False, "RTC設定 OK" & RecvDataLE)


        'RTC確認 -------------------------------------
        'RTC確認 -------------------------------------
        Call WriteLog(False, "RTC確認")


        Call WriteLog(False, "読込開始待ち・・・")

        Call WaitStart(3)

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
            For i = 17 To 42 Step 2
                RTCTime = RTCTime & Chr(CInt("&H" & Mid(RecvDataLE, i, 2)))
            Next i

            PCTime = Format(DateTime.Now, "yy/MM/dd") & Format(DateTime.Now, "HH:mm")

            If RTCTime = PCTime Then Exit Do

            Call WriteLog(False, "RTC確認1回目失敗 SP:" & RTCTime & " PC:" & PCTime)

            If j >= 1 Then
                mErrorComment = "RTC確認2回目失敗 SP:" & RTCTime & " PC:" & PCTime
                Throw New ApplicationException(mErrorComment)
            End If

            j = j + 1
            Call WriteLog(False, "2回目読込開始待ち・・・2秒")

            Call WaitTimer(2)

            DoEvents()

        Loop

        'ResultData.RTC_Read = RecvDataLE
        Call WriteLog(False, "RTC読込 OK : " & RTCTime & " PCTime:" & PCTime)


        ' ---------------------------------------------------------------------------------------


    End Sub

    Public Sub I36(ByVal iName As String)
        '==============================================================================================================
        ' Function : I36:ﾃﾞｰﾀ書込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim mSetting As String = ""
        Dim mDelim As String = ""
        Dim mReadLine As String = ""

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'Test
        ' ---------------------------------------------------------------------------------------
        Call WriteLog(False, "ﾃﾞｰﾀ書込中・・・")

        Dim WriteMACAdres As String
        Dim WriteSerialAdres As String
        Dim HardRev As String

        Dim FuncName As String
        Dim sMac(1) As String
        Dim sSer(2) As String
        Dim AscSer As String
        Dim sDate(2) As String
        Dim AscDay As String
        Dim sDay As String
        Dim sHRev As String
        Dim i As Integer
        Dim rcv As String
        Dim zeroForm As String

        Dim sMacR(1) As String
        Dim sMacAll As String
        Dim sSerR(2) As String
        Dim sSerAll As String
        Dim chSer As String
        Dim sDayR(2) As String
        Dim sDayAll As String
        Dim chDay As String
        Dim sHRevR As String
        Dim chRev As String

        'データ作成 ---------------------------------------------------
        Call WriteLog(False, "ﾃﾞｰﾀ作成")

        'MAC
        WriteMACAdres = IA1()

        sMac(0) = Mid(WriteMACAdres, 1, 8)
        sMac(1) = Mid(WriteMACAdres, 9, 4) & "0000"

        Call WriteLog(False, "MACｱﾄﾞﾚｽ：" & sMac(0) & sMac(1))

        'SER
        AscSer = ""
        'WriteSerialAdres = IA0()
        WriteSerialAdres = SerialRead()
        For i = 1 To Len(WriteSerialAdres)
            AscSer = AscSer & Hex(Asc(Mid(WriteSerialAdres, i, 1)))
        Next
        sSer(0) = Mid(AscSer, 1, 8)
        sSer(1) = Mid(AscSer, 9, 8)
        sSer(2) = Mid(AscSer, 17, 2) & "000000"     '"303000"

        Call WriteLog(False, "ｼﾘｱﾙNo.：" & sSer(0) & sSer(1) & sSer(2))

        'Day
        AscDay = ""
        sDay = Format(DateTime.Now, "yyyy/MM/dd")
        For i = 1 To Len(sDay)
            AscDay = AscDay & Hex(Asc(Mid(sDay, i, 1)))
        Next
        sDate(0) = Mid(AscDay, 1, 8)
        sDate(1) = Mid(AscDay, 9, 8)
        sDate(2) = Mid(AscDay, 17, 4) & "0000"      '"3030"

        Call WriteLog(False, "日付：" & sDate(0) & sDate(1) & sDate(2))

        'HardRev
        zeroForm = ""
        HardRev = HardRevRead()
        sHRev = Hex(CInt(HardRev))
        For i = 1 To 4 - Len(sHRev)
            zeroForm = zeroForm & "0"
        Next
        sHRev = zeroForm & sHRev & "0000"

        Call WriteLog(False, "HardRev：" & sHRev)


        '書き込み開始 --------------------------------------------------
        FuncName = "セクタイレース"
        Call WriteLog(False, FuncName & "中・・・約２秒待ち")

        rcv = CommandComm("TS,22,0,0,107C0000,107C003F,0,", "Test", 0)
        If rcv <> "TS,22,0,0,0,107C0000," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        ' ------------------------------------------------------
        FuncName = "MAC書き込み-1/2"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sMac(0))

        rcv = CommandComm("TS,23,0,0,107C0010," & sMac(0) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0010," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "MAC書き込み-2/2"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sMac(1))

        rcv = CommandComm("TS,23,0,0,107C0014," & sMac(1) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0014," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        ' ------------------------------------------------------
        FuncName = "SER書き込み-1/3"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sSer(0))

        rcv = CommandComm("TS,23,0,0,107C0000," & sSer(0) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0000," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "SER書き込み-2/3"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sSer(1))

        rcv = CommandComm("TS,23,0,0,107C0004," & sSer(1) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0004," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "SER書き込み-3/3"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sSer(2))

        rcv = CommandComm("TS,23,0,0,107C0008," & sSer(2) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0008," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        ' ------------------------------------------------------
        FuncName = "Date書き込み-1/3"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sDate(0))

        rcv = CommandComm("TS,23,0,0,107C0020," & sDate(0) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0020," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "Date書き込み-2/3"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sDate(1))

        rcv = CommandComm("TS,23,0,0,107C0024," & sDate(1) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0024," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "Date書き込み-3/3"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sDate(2))

        rcv = CommandComm("TS,23,0,0,107C0028," & sDate(2) & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0028," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        ' ------------------------------------------------------
        FuncName = "H-Rev書き込み"
        Call WriteLog(False, FuncName & "中・・・")
        Call WriteLog(False, "書き込みデータ:" & sHRev)

        rcv = CommandComm("TS,23,0,0,107C0030," & sHRev & ",0,", "Test", 0)
        If rcv <> "TS,23,0,0,0,107C0030," Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)
        Call WriteLog(False, "")

        '読み返し -----------------------------------------------------------------
        Call WriteLog(False, "読み返し確認")
        Call WriteLog(False, "")

        ' ------------------------------------------------------
        FuncName = "MAC読み込み-1/2"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0010,4,", "Test", 0)
        sMacR(0) = "0000" & Mid(rcv, 14, 4)
        If sMacR(0) <> sMac(0) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "MAC読み込み-2/2"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0014,4,", "Test", 0)
        sMacR(1) = Mid(rcv, 14, 8)
        If sMacR(1) <> sMac(1) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '書き込み判定
        sMacAll = sMacR(0) & Strings.Left(sMacR(1), 4)
        If WriteMACAdres <> sMacAll Then
            mErrorComment = "MACｱﾄﾞﾚｽ読み返し　異常 " & sMacAll
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "MACｱﾄﾞﾚｽ読み返し正常：" & sMacAll)
        Call WriteLog(False, "")

        ' ------------------------------------------------------
        FuncName = "SER読み込み-1/3"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0000,4,", "Test", 0)
        sSerR(0) = Mid(rcv, 14, 8)
        If sSerR(0) <> sSer(0) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "SER読み込み-2/3"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0004,4,", "Test", 0)
        sSerR(1) = Mid(rcv, 14, 8)
        If sSerR(1) <> sSer(1) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "SER読み込み-3/3"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0008,4,", "Test", 0)
        sSerR(2) = Mid(rcv, 14, 8)
        If sSerR(2) <> sSer(2) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '書き込み判定
        sSerAll = sSerR(0) & sSerR(1) & sSerR(2)
        chSer = ""
        For i = 1 To Len(sSerAll) - 6 Step 2
            chSer = chSer & Chr(CInt("&H" & Mid(sSerAll, i, 2)))
        Next

        If WriteSerialAdres <> chSer Then
            mErrorComment = "ｼﾘｱﾙﾅﾝﾊﾞ読み返し　異常 " & chSer
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "ｼﾘｱﾙﾅﾝﾊﾞ読み返し正常：" & chSer)
        Call WriteLog(False, "")

        ' ------------------------------------------------------
        FuncName = "Date読み込み-1/3"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0020,4,", "Test", 0)
        sDayR(0) = Mid(rcv, 14, 8)
        If sDayR(0) <> sDate(0) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "Date読み込み-2/3"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0024,4,", "Test", 0)
        sDayR(1) = Mid(rcv, 14, 8)
        If sDayR(1) <> sDate(1) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        FuncName = "Date読み込み-3/3"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0028,4,", "Test", 0)
        sDayR(2) = Mid(rcv, 14, 8)
        If sDayR(2) <> sDate(2) Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '書き込み判定
        sDayAll = sDayR(0) & sDayR(1) & sDayR(2)
        chDay = ""
        For i = 1 To Len(sDayAll) - 4 Step 2
            chDay = chDay & Chr(CInt("&H" & Mid(sDayAll, i, 2)))
        Next

        If Format(DateTime.Now, "yyyy/MM/dd") <> chDay Then
            mErrorComment = "日付読み返し　異常 " & chDay
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "日付読み返し正常：" & chDay)
        Call WriteLog(False, "")

        ' ------------------------------------------------------
        FuncName = "H-Rev読み込み"
        Call WriteLog(False, FuncName & "中・・・")

        rcv = CommandComm("TS,6,0,0,107C0030,4,", "Test", 0)
        sHRevR = Mid(rcv, 14, Len(rcv) - 23)

        zeroForm = ""
        For i = 1 To 8 - Len(sHRevR)
            zeroForm = zeroForm & "0"
        Next
        sHRevR = zeroForm & sHRevR

        If sHRevR <> sHRev Then
            mErrorComment = FuncName & "　異常 " & rcv
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, FuncName & " OK " & rcv)

        '書き込み判定
        chRev = Strings.Left(sHRevR, 4)
        chRev = CStr(Val("&H" & chRev))
        If Val(HardRev) <> Val(chRev) Then
            mErrorComment = "H-Rev読み返し　異常 " & "Rev=" & chRev
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog(False, "H-Rev読み返し正常：" & "Rev=" & chRev)

        ' ---------------------------------------------------------------------------------------

    End Sub


    Public Sub I38(ByVal iName As String)
        '==============================================================================================================
        ' Function : I38:SDｶｰﾄﾞ検査-ﾉﾝﾗｲﾄﾌﾟﾛﾃｸﾄ
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

        'Test
        ' ---------------------------------------------------------------------------------------
        Dim rcv As String

        'WriteProtect無し確認 -------------------------
        Call WriteLog(False, "WriteProtect無し確認・・・約２秒待ち")

        Try
            rcv = CommandComm("TS,6,0,0,0021C224,4,", "Test", 5)
            '    If rcv <> "TS,6,0,0,0,4,A393FFF8,21C224," Then
            If Mid(rcv, 14, 4) <> "A393" Then
                mErrorComment = "WriteProtect無し確認　異常 " & rcv
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "WriteProtect無し確認 OK " & rcv)

    End Sub

    Public Sub I40(ByVal iName As String)
        '==============================================================================================================
        ' Function : I40:SP69ﾌｧｰﾑ書込み
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Call WriteLog(True, "SP69ﾌｧｰﾑ書込み")

        'メッセージ表示 -------------------------------------------------
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\SP69WriteCheck.ppt"
            dfrm.DISP_TITLENAME = iName
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"

            Call dfrm.frmShow()

            Do
                'wait push button -------------------
                If dfrm.DIALOG_RESULT = "NG" Then
                    mErrorComment = "中止が押されました。"
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

    End Sub


    Public Sub I42(ByVal iName As String)
        '==============================================================================================================
        ' Function : I42:SP69形名照合
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim recData As String = ""
        Dim recDataStr As String = ""
        Dim readModel As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'SP COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = Chr(3)
        cISPComm.PortSetup()

        '形名読み出し
        Try
            Call WriteLog(False, "Send : " & "01021WRDZ00121,8")
            cISPComm.SendCmd = Chr(2) & "01021WRDZ00121,8" & Chr(3)
            Call cISPComm.FuncSendRecvCmd()
            recData = cISPComm.RECEIVE_DATA : recData &= Chr(3)
            If InStr(recData, "OK") = 0 Then            'OKが返ってこないときエラー
                mErrorComment = "返信データ異常：OKなし"
                Throw New ApplicationException(mErrorComment)
            Else
                'OK以前とETXを除く
                recData = Mid(recData, InStr(recData, "OK") + 2, _
                            InStr(recData, Chr(3)) - (InStr(recData, "OK") + 2))
                Call WriteLog(False, "Recv : " & recData)
                For i = 1 To Len(recData) Step 2
                    recDataStr = recDataStr & Chr(CInt("&H" & Mid(recData, i, 2)))
                Next i
                Call WriteLog(False, "Recv : " & recDataStr)
            End If
            readModel = Strings.Left(recDataStr, 8)
            Call WriteLog(False, "形名読込み:" & readModel)

            If readModel <> "F3SP699S" Then
                mErrorComment = "形名照合ｴﾗｰ：" & readModel
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

    End Sub

    Public Sub I43(ByVal iName As String)
        '==============================================================================================================
        ' Function : I43:主銘板ﾃﾞｰﾀ取得
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim recData As String = ""
        Dim recDataStr As String = ""

        Dim readSerial As String
        Dim readDate As String
        Dim readMACid As String
        'Dim readHRev As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Call WriteLog(True, "主銘板ﾃﾞｰﾀ取得")

        'SP COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = Chr(3)
        cISPComm.PortSetup()

        'シリアル番号読み出し
        Try
            Call WriteLog(False, "Send : " & "01021WRDZ00513,5")
            cISPComm.SendCmd = Chr(2) & "01021WRDZ00513,5" & Chr(3)
            Call cISPComm.FuncSendRecvCmd()
            recData = cISPComm.RECEIVE_DATA : recData &= Chr(3)
            If InStr(recData, "OK") = 0 Then            'OKが返ってこないときエラー
                mErrorComment = "返信データ異常：OKなし"
                Throw New ApplicationException(mErrorComment)
            Else
                'OK以前とETXを除く
                recData = Mid(recData, InStr(recData, "OK") + 2, _
                            InStr(recData, Chr(3)) - (InStr(recData, "OK") + 2))
                Call WriteLog(False, "Recv : " & recData)
                For i = 1 To Len(recData) Step 2
                    recDataStr = recDataStr & Chr(CInt("&H" & Mid(recData, i, 2)))
                Next i
                Call WriteLog(False, "Recv : " & recDataStr)
            End If
            readSerial = Strings.Left(recDataStr, 9)
            Call WriteData("LineSerial", readSerial)
            Call WriteLog(False, "シリアル番号読込みOK:" & readSerial)
        Catch ex As Exception
            Call cISPComm.PortClose()
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        '検査日読み出し
        recData = ""
        recDataStr = ""
        Try
            cISPComm.SendCmd = Chr(2) & "01021WRDZ00521,7" & Chr(3)
            Call cISPComm.FuncSendRecvCmd()
            recData = cISPComm.RECEIVE_DATA : recData &= Chr(3)
            If InStr(recData, "OK") = 0 Then            'OKが返ってこないときエラー
                mErrorComment = "返信データ異常：OKなし"
                Throw New ApplicationException(mErrorComment)
            Else
                'OK以前とETXを除く
                recData = Mid(recData, InStr(recData, "OK") + 2, _
                            InStr(recData, Chr(3)) - (InStr(recData, "OK") + 2))
                Call WriteLog(False, "Recv : " & recData)
                For i = 1 To Len(recData) Step 2
                    recDataStr = recDataStr & Chr(CInt("&H" & Mid(recData, i, 2)))
                Next i
                Call WriteLog(False, "Recv : " & recDataStr)
            End If
            readDate = Strings.Left(recDataStr, 10)
            Call WriteData("MacAddress2", "検査日：" & readDate)
            Call WriteLog(False, "検査日読込みOK:" & readDate)
        Catch ex As Exception
            Call cISPComm.PortClose()
            cISPComm = Nothing
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        'MACアドレス読み出し
        recData = ""
        recDataStr = ""
        Try
            cISPComm.SendCmd = Chr(2) & "01021WRDZ00114,2" & Chr(3)
            Call cISPComm.FuncSendRecvCmd()
            recData = cISPComm.RECEIVE_DATA : recData &= Chr(3)
            If InStr(recData, "OK") = 0 Then            'OKが返ってこないときエラー
                mErrorComment = "返信データ異常：OKなし"
                Throw New ApplicationException(mErrorComment)
            Else
                'OK以前とETXを除く
                recData = Mid(recData, InStr(recData, "OK") + 2, _
                            InStr(recData, Chr(3)) - (InStr(recData, "OK") + 2))
                Call WriteLog(False, "Recv : " & recData)
            End If
            readMACid = "0000" & Mid(recData, 5, 4) & Strings.Left(recData, 4)
            Call WriteData("MacAddress", readMACid)
            Call WriteLog(False, "MACid読込みOK:" & readMACid)
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

    End Sub

    Public Sub I44(ByVal iName As String)
        '==============================================================================================================
        ' Function : I44:工場出荷設定
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim cEI As cls130_EI_GAZO
        Dim fCam As Boolean = False
        Dim res As MsgBoxResult

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================

        'ｽﾛｯﾄ5で電源ON
        Call WriteLog(True, "ｽﾛｯﾄ5で電源ON")
        Call I90("ｽﾛｯﾄ5")

        'ｶﾒﾗでRDY点灯までｳｪｲﾄ                           ﾀｲﾑｱｳﾄ30秒
        cEI = New cls130_EI_GAZO
        Try
            cEI.IPAdd = "192.168.1.30" : cEI.MasterFile = "F3SP76-7N_LED.jpg"
            If cEI.MasterFileRead = False Then
                fCam = False
                Throw New ApplicationException(mErrorComment)
            Else
                fCam = True
            End If
        Catch ex As Exception
            cEI = Nothing
        End Try

        Select Case fCam
            Case True       'カメラありのとき
                wTime = System.Environment.TickCount
                Try
                    '1回画像測定
                    Call cEI.GAZO_TEST()
                    If Mid(cEI.RESULT, 1, 1) = "0" Then     'RDYが点いているように見えた場合
                        res = MsgBox("電源ON直後にRDYが点灯しているようです。" & vbCrLf & _
                                     "MODEｽｲｯﾁがCにｾｯﾄされていますか？15～30秒間1,2,4,8が点滅していましたか？" & vbCrLf & _
                                     "", MsgBoxStyle.YesNo)
                        If res = MsgBoxResult.No Then
                            mErrorComment = "SP6x初期化異常：MODEｽｲｯﾁ設定異常"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    End If
                    Do
                        Call cEI.GAZO_TEST()
                        Call WaitTimer(1)
                        If (System.Environment.TickCount - wTime) / 1000 > 30 Then
                            res = MsgBox("RDYが点灯していますか？", MsgBoxStyle.YesNo)
                            If res = MsgBoxResult.No Then
                                mErrorComment = "SP6x初期化異常：RDY点灯確認できません。"
                                Throw New ApplicationException(mErrorComment)
                            Else
                                Exit Do
                            End If
                        End If
                    Loop Until cEI.RESULT = "0,1,1,1,1,1,1,1,1,1,1,1"
                Catch ex As Exception
                    cEI = Nothing
                    Call WriteLog(False, mErrorComment)
                    Call I91("")    'ﾓｼﾞｭｰﾙ電源OFF
                    Throw New ApplicationException(mErrorComment)
                End Try
            Case False      'カメラなしのとき
                Try
                    res = MsgBox("ｶﾒﾗ無し。RDY点灯するまで待ってください。RDY点灯したらOKを押してください。", MsgBoxStyle.YesNo)
                    If res = MsgBoxResult.No Then
                        mErrorComment = "SP6x初期化異常：RDY点灯確認できません。"
                        Throw New ApplicationException(mErrorComment)
                    End If
                Catch ex As Exception
                    Call WriteLog(False, mErrorComment)
                    Call I91("")    'ﾓｼﾞｭｰﾙ電源OFF
                    Throw New ApplicationException(mErrorComment)
                End Try
        End Select

        Call WaitStart(3)
        Call WriteLog(False, "初期化正常")

    End Sub

    Public Sub I45(ByVal iName As String)
        '==============================================================================================================
        ' Function : I45:ﾛｰﾀﾘｰｽｲｯﾁ位置0確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rcv As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Call WriteLog(False, "ﾛｰﾀﾘｽｲｯﾁ　0変更メッセージ")
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\RotarySW0.ppt"
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

        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1"
        cISPComm.TERMINAL_CHAR = vbCrLf
        cISPComm.PortSetup()

        '検査開始
        cISPComm.SendCmd = "TS,6,0,0,0021C324,4,"
        Try
            Call cISPComm.FuncSendRecvCmd()
            rcv = cISPComm.RECEIVE_DATA
            'ﾛｰﾀﾘSW設定確認 0以外はNG
            If Mid(rcv, 14, 4) <> "FFFF" Then
                mErrorComment = "ﾛｰﾀﾘｰｽｲｯﾁが0になっていません。"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "ﾛｰﾀﾘｰｽｲｯﾁ設定0確認OK")
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

    End Sub

    Public Sub I46(ByVal iName As String)
        '==============================================================================================================
        ' Function : I46:ﾛｰﾀﾘｰｽｲｯﾁ位置C確認
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================

        Dim wTime As Double = 0
        Dim rcv As String

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Call WriteLog(False, "ﾛｰﾀﾘｽｲｯﾁ　C変更メッセージ")
        Try
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\RotarySWC.ppt"
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

        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1"
        cISPComm.TERMINAL_CHAR = vbCrLf
        cISPComm.PortSetup()

        '検査開始
        cISPComm.SendCmd = "TS,6,0,0,0021C324,4,"
        Try
            Call cISPComm.FuncSendRecvCmd()
            rcv = cISPComm.RECEIVE_DATA
            'ﾛｰﾀﾘSW設定確認 C以外はNG
            If Mid(rcv, 14, 4) <> "F9FF" Then
                mErrorComment = "ﾛｰﾀﾘｰｽｲｯﾁがCになっていません。"
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "ﾛｰﾀﾘｰｽｲｯﾁ設定C確認OK")
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
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

    Public Sub I60(ByVal iName As String)
        '==============================================================================================================
        ' Function : I60:XIRQ7,4ｱｻｰﾄﾃｽﾄ
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")
        '==========================================================================================
        Dim rcv As String

        'XPFAILをｱｻｰﾄしてXIRQ7がｱｻｰﾄされることを確認
        Call WriteLog(False, "XPFAILをｱｻｰﾄしてXIRQ7がｱｻｰﾄされることを確認")
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call WriteLog(False, "ﾒｲﾝCPUのXPFAILﾏｽｸ")
        '事前準備1
        Call cMcm.MEM_WRITE(1, "38800013", "00")
        '事前準備2
        Call cMcm.MEM_WRITE(1, "38800011", "00")
        '割り込み禁止設定
        rcv = CommandComm("TS,7,0,0,0021C330,4,0,", "Test", 5)

        Try
            'XPFAILｾｯﾄ
            If cBUS.IOBusTstSetCpuXPFAIL() = False Then
                Call cBUS.IOBusTstReSetCpuXPFAIL()
                mErrorComment = "XPFAILｾｯﾄ異常"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Call cMcm.MEM_WRITE(1, "38800011", "80")
            cMcm = Nothing
            cBUS = Nothing
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog(False, "XPFAILｾｯﾄOK")

        Try
            'ﾃｽﾄｺﾏﾝﾄﾞでXIRQ7(0x21C324)をﾘｰﾄﾞ
            Call WriteLog(False, "ﾃｽﾄｺﾏﾝﾄﾞでXIRQ7(0x21C324)をﾘｰﾄﾞ")
            rcv = CommandComm("TS,6,0,0,0021C324,4,", "Test", 5)
            If Mid(rcv, 14, 1) <> "B" Then Throw New ApplicationException(mErrorComment)
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XIRQ7確認異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXPFAIL()
            Call WriteLog(False, "XPFAILﾘｾｯﾄOK")
            Call WriteLog(False, "ﾒｲﾝCPUのXPFAILﾏｽｸ解除")
            Call cMcm.MEM_WRITE(1, "38800011", "80")
            cMcm = Nothing
            cBUS = Nothing
            '割り込み設定
            rcv = CommandComm("TS,7,0,0,0021C330,4,40000000,", "Test", 5)
        End Try

        Call WriteLog(False, "XIRQ7：OK")

        'XIRQ3をｱｻｰﾄしてXIRQ4がｱｻｰﾄされることを確認
        Call WriteLog(False, "XIRQ3をｱｻｰﾄしてXIRQ4がｱｻｰﾄされることを確認")
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = IP_U2                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'ﾒｲﾝCPUのXIRQﾏｽｸ
        Call WriteLog(False, "ﾒｲﾝCPUのXIRQﾏｽｸ")
        Call cMcm.MEM_WRITE(1, "3880021E", "00")
        'KP側IOﾊﾞｽ割り込みｲﾈｰﾌﾞﾙ
        rcv = CommandComm("TS,7,0,0,1660B21E,1,F0,", "Test", 5)
        'CPU側割り込みﾏｽｸ
        rcv = CommandComm("TS,7,0,0,0021C230,4,0,", "Test", 5)

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Try
            'XIRQ3をｱｻｰﾄ
            Call WriteLog(False, "XIRQ3をｱｻｰﾄ")
            Call cBUS.BUSMONITOR_TB04_CPU_XIRQChange(3)

            'ﾃｽﾄｺﾏﾝﾄﾞでXIRQ4(0x21C224)をﾘｰﾄﾞ
            rcv = CommandComm("TS,6,0,0,0021C224,4,", "Test", 5)
            If Mid(rcv, 14, 1) <> "2" Then Throw New ApplicationException(mErrorComment)

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XIRQ異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.BUSMONITOR_TB04_CPU_XIRQChange(5)
            'ﾒｲﾝCPUのXIRQﾏｽｸ解除
            Call WriteLog(False, "ﾒｲﾝCPUのXIRQﾏｽｸ解除")
            Call cMcm.MEM_WRITE(1, "3880021E", "88")
            cBUS = Nothing
            cMcm = Nothing
            'KP側IOﾊﾞｽ割り込みﾃﾞｨｾｰﾌﾞﾙ
            rcv = CommandComm("TS,7,0,0,1660B21E,1,0,", "Test", 5)

        End Try

        Call WriteLog(False, "XIRQ4：OK")


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

        Dim scmd As String
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
            cLnk = New cls080_F3LinkToLE
            cLnk.IPAdd = IP_U1
            If InStr(iName, "BOOT") <> 0 Then
                scmd = "01" & "BWRY" & "902" & ",001,1"
            Else
                scmd = "01" & "BWRY" & "902" & ",001,0"
            End If
            Call cLnk.F3PCLinkCommunication(scmd, rCmd)
            If cLnk.ERR_COMMENT <> "" Then
                mErrorComment = cLnk.ERR_COMMENT
                Throw New ApplicationException(mErrorComment)
            End If
            cBUS = New cls122_FAM3_IOBUS_Ctrl
            Call cBUS.BUSMONITOR_TB04_CPU_START(slot)

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            cLnk = Nothing
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
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6X\START.ppt"
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

        Try

            'Setting message =======================================
            dfrm = New frm500_MessagePopUp
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6x\PassWin.ppt"
            'cSys.SignalTower("GREEN", "OFF", "ON", "OFF")
            dfrm.LED_COLOR = "GREEN" : dfrm.BUZZ1 = "OFF" : dfrm.BUZZ2 = "ON" : dfrm.FLASH = "ON"
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

        Call WriteLog(False, "MACｱﾄﾞﾚｽ :" & mMACAddress)
        Call WriteLog(False, "Lineｼﾘｱﾙ :" & mLineSerial)

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

    Private Function IA1() As String
        '==============================================================================================================
        ' Function : IA1:Mac Address Port
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        IA1 = ""
        Dim cSP7x As cls2D2_F3SP7x_Inspection
        cSP7x = New cls2D2_F3SP7x_Inspection
        Try
            IA1 = cSP7x.IA1()
        Catch ex As Exception
            mErrorComment = cSP7x.ERR_COMMENT
            Throw New ApplicationException(mErrorComment)
        Finally
            cSP7x = Nothing
        End Try

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
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\F3SP6x\PassWin.ppt"
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
            cLblPrnt.PrintForm = "SP6x"
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
        ' Function : Write Value data(F3SP6XData.mdb)
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

        dbName = mAppPath & "210 Result Data\F3SP6XData.mdb"
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

    Public Function TestModeStart() As Boolean

        TestModeStart = False

        'SP TestMode COM OPEN
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,E,8,1" : cISPComm.TERMINAL_CHAR = vbCrLf
        Call cISPComm.PortSetup()

        Dim send(19) As Byte
        Dim i As Integer

        '送信データの作成 (STX SP TEST MODE START ETX のbyte送信)
        send(0) = &H2
        send(19) = &H3

        For i = 1 To 18
            send(i) = CByte(Asc(Mid("SP TEST MODE START", i, 1)))
        Next

        cISPComm.SendByteData = send

        Try
            Call cISPComm.FuncSendByteData()
        Catch ex As Exception
            mErrorComment = "ﾃｽﾄﾓｰﾄﾞ送信　異常"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

        Call WriteLog(False, "TestMode移行　5秒待ち")
        Call WaitStart(5)

        TestModeStart = True

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

    Private Function FTPConnect(ByVal DirName As String) As Boolean
        '==========================================================================================
        ' 関数名 ： FTPConnect
        ' 機　能 ： FTP接続→ｶﾚﾝﾄﾃﾞｨﾚｸﾄﾘ指定
        ' 引　数 ： DirName
        ' 戻り値 ： 成功/失敗
        ' 備　考 ：
        '==========================================================================================

        'Dim sRet As String

        FTPConnect = False

        Try
            'ファイル一覧を取得するディレクトリのURI
            Dim u As New Uri("ftp://192.168.0.2/" & DirName)

            'FtpWebRequestの作成
            Dim ftpReq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
            'ログインユーザー名とパスワードを設定
            ftpReq.Credentials = New System.Net.NetworkCredential("", "")
            'MethodにWebRequestMethods.Ftp.PrintWorkingDirectoryを設定
            ftpReq.Method = System.Net.WebRequestMethods.Ftp.PrintWorkingDirectory

            'FtpWebResponseを取得
            Dim ftpRes As System.Net.FtpWebResponse = CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)

            If InStr(ftpRes.StatusDescription.ToString, DirName) = 0 Then
                mErrorComment = "Current Dir Err " & DirName
                Throw New ApplicationException(mErrorComment)
            End If

            '閉じる
            ftpRes.Close()

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "FTP Connect Error"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        FTPConnect = True

    End Function

    Public Function ReWriteSD() As Boolean
        '==========================================================================================
        ' 関数名 ： ReWriteSD
        ' 機　能 ： 100M通信検査時にSD内ﾗﾀﾞｰを書換え
        ' 引　数 ： DirName
        ' 戻り値 ： 成功/失敗
        ' 備　考 ：
        '==========================================================================================

        'Dim DirName As String
        Dim DelFile As String = ""
        'Dim sRet As String
        Dim strName As String
        Dim rFileName() As String = {""}
        Dim i As Integer
        Dim PutFile As String
        Dim szFileRemote As String
        Dim szFileLocal As String = ""
        Dim Model As String

        ReWriteSD = False

        'Connect -------------------------------
        'ファイル一覧を取得するディレクトリのURI
        Dim u As New Uri("ftp://192.168.0.2//CARD1/PROJECT")

        'FtpWebRequestの作成
        Dim ftpReq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
        'ログインユーザー名とパスワードを設定
        ftpReq.Credentials = New System.Net.NetworkCredential("", "")
        'MethodにWebRequestMethods.Ftp.ListDirectoryDetails("LIST")を設定
        ftpReq.Method = System.Net.WebRequestMethods.Ftp.ListDirectory
        '要求の完了後に接続を閉じる
        ftpReq.KeepAlive = False
        'PASSIVEモードを無効にする
        ftpReq.UsePassive = False

        'FtpWebResponseを取得
        Dim ftpRes As System.Net.FtpWebResponse = _
            CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
        'FTPサーバーから送信されたデータを取得
        Dim sr As New System.IO.StreamReader(ftpRes.GetResponseStream())
        Dim res As String = sr.ReadToEnd()
        rFileName = Split(res, vbCrLf)
        'ファイル一覧を表示
        sr.Close()
        '閉じる
        ftpRes.Close()

        'Next Serch ----------------------
        i = 0
        Do
            If rFileName(i) = "" Then Exit Do
            Try
                If InStr(rFileName(i), ",") <> 0 Then
                    strName = Strings.Left(rFileName(i), InStr(rFileName(i), ",") - 1)
                Else
                    strName = rFileName(i)
                End If
                'If Strings.Left(strName, 3) = "SP6" Then DelFile = strName 'v0.07でｺﾒﾝﾄ
                If Strings.Left(strName, 3) = "SP6" Or Strings.Left(strName, 3) = "SP7" Then DelFile = strName
                If Strings.Left(strName, 3) = "SPV" Then DelFile = strName 'Rev0.12ADD
            Catch ex As Exception

            End Try
            i = i + 1
            DoEvents()
        Loop Until i > UBound(rFileName)

        'If i > 2 Or i < 1 Then
        If i < 1 Then
            MsgBox("SDカードの内容を確認して下さい。File数" & CStr(i), vbCritical)
            Exit Function
        End If

        'ファイルがある場合に削除 ----------------------
        If DelFile <> "" Then
            If DeleteFTPFile(DelFile) = False Then
                MsgBox("SDファイル削除失敗", vbCritical)
                Exit Function
            End If
        End If

        Model = ModelRead()
        'SDにファイル出力 ------------------------------
        Select Case Mid(Model, 3, 4)    'Ver0.01にて追加
            Case "SP69"
                PutFile = "SP67-6S" & ".YPJC"
            Case Else
                PutFile = Mid(Model, 3, 7) & ".YPJC"
        End Select

        'v0.07
        Select Case Mid(Model, 3, 3)
            Case "SP6"
                szFileLocal = mAppPath & "600 Equipment Program\SP6X\TestProgram\" & PutFile
            Case "SP7"
                szFileLocal = mAppPath & "600 Equipment Program\SP7X\TestProgram\" & PutFile
            Case "SPV"                                                                                 'Rev0.12ADD
                szFileLocal = mAppPath & "600 Equipment Program\SPVX\TestProgram\" & PutFile           'Rev0.12ADD
        End Select

        'szFileLocal = mAppPath & "600 Equipment Program\SP6X\TestProgram\" & PutFile
        szFileRemote = PutFile

        If PutFTPFile(szFileLocal, szFileRemote) = False Then
            MsgBox("FTP出力失敗")
            Exit Function
        End If

        ReWriteSD = True

    End Function

    Private Function PutFTPFile(ByVal szFileLocal As String, ByVal szFileRemote As String) As Boolean
        '==========================================================================================
        ' 関数名 ： PutPFile
        ' 機　能 ： 指定ファイルのコピー
        ' 引　数 ： ファイル名
        ' 戻り値 ： 成功/失敗
        ' 備　考 ：
        '==========================================================================================

        PutFTPFile = False

        'アップロードするファイル
        Dim upFile As String = szFileLocal
        'アップロード先のURI
        Dim u As New Uri("ftp://192.168.0.2//CARD1/PROJECT/" & szFileRemote)

        Try
            'FtpWebRequestの作成
            Dim ftpReq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
            'ログインユーザー名とパスワードを設定
            ftpReq.Credentials = New System.Net.NetworkCredential("", "")
            'MethodにWebRequestMethods.Ftp.UploadFile("STOR")を設定
            ftpReq.Method = System.Net.WebRequestMethods.Ftp.UploadFile
            '要求の完了後に接続を閉じる
            ftpReq.KeepAlive = False
            'ASCIIモードで転送する
            ftpReq.UseBinary = False
            'PASVモードを無効にする
            ftpReq.UsePassive = False

            'ファイルをアップロードするためのStreamを取得
            Dim reqStrm As System.IO.Stream = ftpReq.GetRequestStream()
            'アップロードするファイルを開く
            Dim fs As New System.IO.FileStream(upFile, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            'アップロードStreamに書き込む
            Dim buffer(1023) As Byte
            While True
                Dim readSize As Integer = fs.Read(buffer, 0, buffer.Length)
                If readSize = 0 Then
                    Exit While
                End If
                reqStrm.Write(buffer, 0, readSize)
            End While
            fs.Close()
            reqStrm.Close()

            'FtpWebResponseを取得
            Dim ftpRes As System.Net.FtpWebResponse = _
                CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
            'FTPサーバーから送信されたステータスを表示
            Console.WriteLine("{0}: {1}", ftpRes.StatusCode, ftpRes.StatusDescription)
            '閉じる
            ftpRes.Close()

        Catch ex As Exception
            Exit Function
        End Try

        PutFTPFile = True

    End Function

    Private Function DeleteFTPFile(ByVal FileName As String) As Boolean
        '==========================================================================================
        ' 関数名 ： DeleteFTPFile
        ' 機　能 ： 指定ファイルの削除
        ' 引　数 ： ファイル名
        ' 戻り値 ： 成功/失敗
        ' 備　考 ：
        '==========================================================================================

        DeleteFTPFile = False

        Try
            '削除するファイルのURI
            Dim u As New Uri("ftp://192.168.0.2//CARD1/PROJECT/" & FileName)

            'FtpWebRequestの作成
            Dim ftpReq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(u), System.Net.FtpWebRequest)
            'ログインユーザー名とパスワードを設定
            ftpReq.Credentials = New System.Net.NetworkCredential("", "")
            'MethodにWebRequestMethods.Ftp.DeleteFile(DELE)を設定
            ftpReq.Method = System.Net.WebRequestMethods.Ftp.DeleteFile
            'FtpWebResponseを取得
            Dim ftpRes As System.Net.FtpWebResponse = CType(ftpReq.GetResponse(), System.Net.FtpWebResponse)
            '閉じる
            ftpRes.Close()
        Catch ex As Exception
            Exit Function
        End Try

        DeleteFTPFile = True

    End Function

    Private Function CommandComm(ByVal SendCommand As String, ByVal MODE As String, ByVal ReadTimeOut As Integer) As String

        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.TERMINAL_CHAR = vbCrLf

        If MODE = "Boot" Then
            cISPComm.SETTING = "9600,N,8,1"
        Else
            cISPComm.SETTING = "9600,E,8,1"
        End If

        Try
            cISPComm.SendCmd = SendCommand
            Call WriteLog(False, "Send : " & cISPComm.SendCmd)
            'Call cISPComm.FuncSendRecvCmd()
            Call cISPComm.PortSetup()
            cISPComm.TIMEOUT = ReadTimeOut
            Call cISPComm.FuncSendCmd()
            Call cISPComm.FuncRecvCmd()
            Call WriteLog(False, "Recv : " & cISPComm.RECEIVE_DATA)
            CommandComm = cISPComm.RECEIVE_DATA
        Catch ex As Exception
            mErrorComment = "ｼﾘｱﾙ通信ｴﾗｰ"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
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

    Private Function SubTest01(ByRef NewDrv As String) As Boolean
        '==========================================================================================
        ' 関数名 ： SubTest01
        ' 機　能 ： USB接続
        ' 引　数 ： 無し
        ' 戻り値 ： 成功/失敗
        ' 備　考 ：
        '==========================================================================================

        Dim strDrive1 As String
        Dim strDrive2 As String
        'Dim rcv As String
        Dim strdrv As String
        Dim drvType As String = ""
        Dim BIT As Long
        Dim i As Integer
        Dim j As Integer
        Dim Drive As String

        SubTest01 = False

        'ドライブ一覧取得 -------------------------------
        strDrive1 = "" : strDrive2 = "" : NewDrv = ""
        BIT = 1

        For Each Drive In IO.Directory.GetLogicalDrives
            strDrive1 = strDrive1 & Strings.Left(Drive, 1)
        Next

        Call WriteLog(False, "現在のドライブ構成：" & strDrive1)
        Call WriteLog(False, "USB接続　開始")

        'TS,81,,,通信 --------------------------------
        '送受信
        cISPComm = New cls101_ComRS232C
        cISPComm.RS_COMM = Insfrm.SerialPort1
        cISPComm.PORT = rs232cSPPort
        cISPComm.SETTING = "9600,N,8,1"
        cISPComm.TERMINAL_CHAR = vbCrLf

        Try
            cISPComm.SendCmd = "TS,81,0,0,"
            Call cISPComm.PortSetup()
            Call cISPComm.FuncSendCmd()

            'Call WaitTimer(1)
            'Call cISPComm.FuncKYOUSEIRecvCmd()

            'rcv = cISPComm.RECEIVE_DATA
            'rcv = Strings.Left(rcv, 1)
            'If rcv <> "TS,81,0,0," Then
            '    mErrorComment = "USBﾓｰﾄﾞ(TS,81)　失敗"
            '    Throw New ApplicationException(mErrorComment)
            'End If
        Catch ex As Exception
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cISPComm.PortClose()
            cISPComm = Nothing
        End Try

        'rcv = CommandComm("TS,81,0,0,", "Boot", 5)
        'If rcv <> "TS,81,0,0," Then
        '    MsgBox("USBﾓｰﾄﾞ(TS,81)　失敗")
        '    Throw New ApplicationException(mErrorComment)
        'End If

        'USB確立後のドライブ確認 -------------------------
        Call WriteLog(False, "USB接続　完了")
        Call WriteLog(False, "増加ドライブの確認中・・・")

        '増加待ち ---------------------------------------
        j = 0
        Do
            strDrive2 = ""
            'BIT = 1
            'Drives = GetLogicalDrives()

            For Each Drive In IO.Directory.GetLogicalDrives
                strDrive2 = strDrive2 & Strings.Left(Drive, 1)
            Next

            j = j + 1
            If j >= 30 Then
                MsgBox("ドライブ増加せず")
                Throw New ApplicationException(mErrorComment)
            End If

            DoEvents()

            Call WaitTimer(1)

            DoEvents()

        Loop Until strDrive1 <> strDrive2

        Call WriteLog(False, "接続後のドライブ構成：" & strDrive2)

        'ドライブ文字の確認 -----------------------------------
        For i = 1 To Len(strDrive2)
            strdrv = Mid(strDrive2, i, 1)
            'strDrive1と比べて、新しい文字があったらNewDrvへ格納
            If InStr(strDrive1, strdrv) = 0 Then NewDrv = NewDrv + strdrv
            DoEvents()
        Next

        '新しい文字が1個以外だったら異常
        If Len(NewDrv) <> 1 Then
            MsgBox("ドライブ増加　複数台：" & NewDrv)
            Throw New ApplicationException(mErrorComment)
        End If

        'ドライブ種類の確認 -----------------------------------
        Dim DriveName As New System.IO.DriveInfo(NewDrv)

        Select Case DriveName.DriveType
            Case IO.DriveType.Unknown : drvType = "判定不能"
            Case IO.DriveType.NoRootDirectory : drvType = "存在しない"
            Case IO.DriveType.Removable : drvType = "ﾘﾑｰﾊﾞﾌﾞﾙ"
            Case IO.DriveType.Fixed : drvType = "固定"
            Case IO.DriveType.Network : drvType = "ﾘﾓｰﾄ"
            Case IO.DriveType.CDRom : drvType = "CD-ROM"
            Case IO.DriveType.Ram : drvType = "RAMﾃﾞｨｽｸ"
        End Select

        Call WriteLog(False, "追加ドライブ　：" & NewDrv)
        Call WriteLog(False, "ドライブタイプ：" & drvType)

        'ドライブのタイプが違っていたら異常
        If drvType <> "ﾘﾑｰﾊﾞﾌﾞﾙ" Then
            MsgBox("ドライブタイプ異常：" & drvType)
            Throw New ApplicationException(mErrorComment)
        End If

        DriveName = Nothing
        SubTest01 = True

        Exit Function

NG_Trap:
        Exit Function
ErrTrap:
    End Function

    Private Function SubTest02(ByVal NewDrv As String) As Boolean
        '==========================================================================================
        ' 関数名 ： SubTest02
        ' 機　能 ： ﾌｧｰﾑのｺﾋﾟｰ
        ' 引　数 ： 無し
        ' 返り値 ： 無し
        ' 備　考 ：
        '==========================================================================================
        Dim Model As String = ModelRead()
        Dim BasePass(3) As String
        Dim CopyPass(3) As String
        Dim i As Integer
        'Dim sf As SHFILEOPSTRUCT

        '検査型式の読込み
        Dim dbName As String = mAppPath & "210 Result Data\Temporary.mdb"
        Dim rDTkey As String = ""
        Dim tbName As String
        Dim rField As String = "MSCODE,ProductSoftRev"
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim ResultData_MsCode As String = ""
        Dim ResultData_SoftRev As String = ""
        Dim ResultData_SoftRevSP67 As String = ""

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName
            sSQL = "SELECT " & rField & " FROM " & tbName
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)
            ResultData_MsCode = sData(0)
            ResultData_SoftRev = sData(1)
        Catch ex As Exception
        Finally
            cADO = Nothing
        End Try

        dbName = mAppPath & "500 Model Information\ModelInformation.mdb"
        rField = "SoftRev"

        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblModelInformation"
            cADO.TB_NAME = tbName
            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model Like 'F3SP67%'"
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_JET(fName, sData)
            ResultData_SoftRevSP67 = sData(0)
        Catch ex As Exception
        Finally
            cADO = Nothing
        End Try

        'ﾌｧｲﾙ定義 --------------------------------------
        'コピー元

        If Strings.Left(ResultData_MsCode, 6) = "F3SP66" Then
            BasePass(0) = mAppPath & "520 Product Firmware\SP6X\" & "S9666US-" & ResultData_SoftRev & ".bin"
        ElseIf Strings.Left(ResultData_MsCode, 6) = "F3SP69" Then
            BasePass(0) = mAppPath & "520 Product Firmware\SP6X\" & "S9667US-" & ResultData_SoftRevSP67 & ".bin"
        Else
            BasePass(0) = mAppPath & "520 Product Firmware\SP6X\" & "S9667US-" & ResultData_SoftRev & ".bin"
        End If
        '    BasePass(1) = App.path & "\" & "Firmware" & "\" & "S9660US-" & ResultData.HardRev & ".BIN"
        BasePass(1) = mAppPath & "520 Product Firmware\SP6X\" & "S9660US-" & "00" & ".BIN"
        '    BasePass(2) = App.path & "\" & "Firmware" & "\" & "SPU_SP67.BIN"
        BasePass(2) = mAppPath & "520 Product Firmware\SP6X\" & "MAC.BIN"

        'ｺﾋﾟｰ先
        If Strings.Left(ResultData_MsCode, 6) = "F3SP66" Then
            CopyPass(0) = "F3SP66.BIN"     'NewDrv & ":\" &
        Else
            CopyPass(0) = "F3SP67.BIN"      'NewDrv & ":\" & "F3SP67.BIN"
        End If
        If Strings.Left(ResultData_MsCode, 6) = "F3SP66" Then
            CopyPass(1) = "F3SP66BT.BIN"        'NewDrv & ":\" &
        Else
            CopyPass(1) = "F3SP67BT.BIN"        'NewDrv & ":\" &
        End If

        CopyPass(2) = "MAC.BIN"     'NewDrv & ":\" &

        '    For i = 0 To 2
        '        FileName = Dir(CopyPass(i))
        '        If FileName <> "" Then Kill CopyPass(i)
        '    Next

        'Call mSecWait(1000)
        Call WaitTimer(1)

        'ｺﾋﾟｰ -------------------------------------------
        cFop = New cls012_File_Operation

        'sf.hwnd = frm020_TestInformation.hwnd
        'sf.wFunc = FO_COPY
        'sf.fFlags = NO_CONFIRM

        For i = 0 To 2
            Call WriteLog(False, "ｺﾋﾟｰ中" & CopyPass(i))

            cFop.ORIGINAL_PATH = BasePass(i)
            cFop.FILE_NAME = CopyPass(i)
            cFop.FOLDER_PATH = NewDrv & ":\"
            cFop.FILE_COPY()
            'sf.pFrom = BasePass(i)
            'sf.pTo = CopyPass(i)

            'Ret = SHFileOperation(sf)
            'If Ret <> 0 Then MsgBox("失敗しました。")

            DoEvents()
        Next

        SubTest02 = True

        Exit Function

NG_Trap:
        Exit Function
ErrTrap:

    End Function


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

End Class
