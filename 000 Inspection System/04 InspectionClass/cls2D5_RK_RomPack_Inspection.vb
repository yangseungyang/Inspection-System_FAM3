Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm020_InspectionStartEnd
Imports System.Text
Imports System.Threading


Public Class cls2D5_RK_RomPack_Inspection
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
    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String
    Dim rs232cSPPort As String = "COM5"

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

        Debug.Print(iName)

        Select Case Mid(iName, 1, 3)
            Case "I01" : Call I01(iName)            'I01:ROM PACK検査

                'Case "IA0" : Call IA0(iName)            'IA0:LineSerial
            Case "IP1" : Call IP1(iName)            'IP1:PrintLabel

            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF
            Case "I92" : Call I92(iName)            'I92:検査開始準備
            Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ
            Case "I94" : Call I94(iName)            'I94:Result Log Save        'Rev 0.26 KSJ ADD
                'Case Else : mErrorComment = "検査プロシージャが定義されていません。"   'Rev0.11DEL
            Case Else : mErrorComment = "검사 프로시져가 정의 되어있지 않습니다."   'Rev0.11ADD

        End Select

    End Sub

    Public Sub I01(ByVal iName As String)
        '==============================================================================================================
        ' Function : I01:ROM PACK検査
        ' Data     : Process name
        ' Return   : Nothing
        ' Remarks  :
        '==============================================================================================================
        Dim SendParam As String() = {""}
        Dim Model As String = ModelRead()
        Dim RKDATA As String = ""
        Dim RecvDATA As String()

        Call WriteLog(False, "")
        Call WriteLog(False, iName & " " & Format(DateTime.Now, "HH:mm:ss"))
        Call WriteLog(False, "*")

        '==============================================================================================================

        'SPテストモード移行
        cSP = New cls121_SP_TEST_MODE
        cSP.PORT = rs232cSPPort
        cSP.RS_COMM = Insfrm.SerialPort1

        WaitTimer(5)

        Call WriteLog(False, "SPテストモード移行中。。。しばらくお待ちください。")
        Try                                             'V3.02
            Call cSP.OldSpTESTMODEStart()
            Call WriteLog(False, cSP.LOG)
        Catch ex As Exception
            mErrorComment = "SPﾃｽﾄﾓｰﾄﾞ移行ｴﾗｰ"
            cSP = Nothing
            Call WriteLog(False, cSP.LOG)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End Try

        WaitTimer(7)
        Call WriteLog(False, "SPテストモード移行OK")

        cSP.ICMD = 9
        cSP.IUNIT = 0
        cSP.ISLOT = 0
        cSP.IPARA = 4

        ReDim SendParam(4)
        SendParam(3) = "1"

        Select Case Mid(Model, 1, 7)
            Case "RK10-0N"
                SendParam(2) = "3" : RKDATA = "222"
            Case "RK30-0N"
                SendParam(2) = "4" : RKDATA = "233"
            Case "RK33-0N"
                SendParam(2) = "5" : RKDATA = "3433"
            Case "RK50-0N"
                SendParam(2) = "6" : RKDATA = "3434"
            Case "RK53-0N"
                SendParam(2) = "6" : RKDATA = "3434"
            Case "RK73-0N"
                SendParam(2) = "7" : RKDATA = "3535"
            Case "RK93-0N"
                SendParam(2) = "8" : RKDATA = "3636"
        End Select

        'TEST01 ADDR/DATA BUS CHECK
        Call WriteLog(False, "TEST01 ADDR/DATA BUS CHECK。。。しばらくお待ちください。")
        DoEvents()
        SendParam(1) = "5"
        SendParam(4) = "0"

        cSP.SendPARA = SendParam

        Call cSP.TestComSendRecv()
        Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
        Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
        If RecvDATA(4) <> "0" Then
            mErrorComment = "TEST01 ADDR/DATA BUS CHECK ERROR"
            cSP = Nothing                             'V3.02
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "TEST01 ADDR/DATA BUS CHECK OK!")

        'TEST02 WRITE/READ TEST (TEST DATA=$55)
        Call WriteLog(False, "TEST02 WRITE/READ TEST (TEST DATA=$55)。。。しばらくお待ちください。")
        DoEvents()
        SendParam(1) = "2"
        SendParam(4) = "55"

        cSP.SendPARA = SendParam

        Call cSP.TestComSendRecv()
        Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
        Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
        If RecvDATA(4) <> "0" Then
            mErrorComment = "TEST02 WRITE/READ TEST (TEST DATA=$55) ERROR"
            cSP = Nothing                             'V3.02
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "TEST02 WRITE/READ TEST (TEST DATA=$55) OK!")

        'TEST03 WRITE/READ TEST (TEST DATA=$AA)
        Call WriteLog(False, "TEST03 WRITE/READ TEST (TEST DATA=$AA)。。。しばらくお待ちください。")
        DoEvents()
        SendParam(1) = "2"
        SendParam(4) = "AA"

        cSP.SendPARA = SendParam

        Call cSP.TestComSendRecv()
        Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
        Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
        If RecvDATA(4) <> "0" Then
            mErrorComment = "TEST03 WRITE/READ TEST (TEST DATA=$AA) ERROR"
            cSP = Nothing                             'V3.02
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "TEST03 WRITE/READ TEST (TEST DATA=$AA) OK!")


        'TEST04 出荷設定
        Call WriteLog(False, "TEST04 出荷設定。。。しばらくお待ちください。")
        SendParam(1) = "4"
        SendParam(4) = "0"
        cSP.SendPARA = SendParam

        Call cSP.TestComSendRecv()
        Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
        Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
        If RecvDATA(4) <> "0" Then
            mErrorComment = "TEST04 出荷設定 ERROR"
            cSP = Nothing                             'V3.02
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "TEST04 出荷設定 OK!")

        'TEST05 出荷設定確認
        Dim RKCHK1, RKCHK2, RKCHK3 As String
        Dim RKCHK4 As String = ""
        Call WriteLog(False, "TEST05 出荷設定確認")
        DoEvents()
        RKCHK1 = Mid(RKDATA, 1, 1)
        RKCHK2 = Mid(RKDATA, 2, 1)
        RKCHK3 = Mid(RKDATA, 3, 1)
        If Len(RKDATA) = 4 Then RKCHK4 = Mid(RKDATA, 4, 1)

        'ROM PACK $02 READ Check
        Call WriteLog(False, "ROM PACK $02 READ Check")
        cSP.ICMD = 6
        SendParam(1) = "200005"
        SendParam(2) = "1"
        cSP.IPARA = 2
        cSP.SendPARA = SendParam

        Call cSP.TestComSendRecv()
        Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
        Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
        Call WriteLog(False, "ROM PACK $02 DATA = " & RecvDATA(5))

        If RecvDATA(5) <> RKCHK1 Then
            mErrorComment = "ROM PACK $02 READ Check ERROR"
            cSP = Nothing                             'V3.02
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "ROM PACK $02 READ Check OK!")

        'ROM PACK $03 READ Check
        Call WriteLog(False, "ROM PACK $03 READ Check")
        cSP.ICMD = 6
        SendParam(1) = "200007"
        SendParam(2) = "1"
        cSP.IPARA = 2
        cSP.SendPARA = SendParam

        Call cSP.TestComSendRecv()
        Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
        Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
        Call WriteLog(False, "ROM PACK $03 DATA = " & RecvDATA(5))
        If RecvDATA(5) <> RKCHK2 Then
            mErrorComment = "ROM PACK $03 READ Check ERROR"
            cSP = Nothing                             'V3.02
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "ROM PACK $03 READ Check OK!")

        'ROM PACK $04 READ Check
        Call WriteLog(False, "ROM PACK $04 READ Check")
        cSP.ICMD = 6
        SendParam(1) = "200009"
        SendParam(2) = "1"
        cSP.IPARA = 2
        cSP.SendPARA = SendParam

        Call cSP.TestComSendRecv()
        Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
        Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
        RecvDATA = Split(cSP.RECEIVE_DATA, ",")
        Call WriteLog(False, "ROM PACK $04 DATA = " & RecvDATA(5))
        If RecvDATA(5) <> RKCHK3 Then
            mErrorComment = "ROM PACK $04 READ Check ERROR"
            cSP = Nothing                             'V3.02
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If
        Call WriteLog(False, "ROM PACK $04 READ Check OK!")

        'ROM PACK $05 READ Check
        If Len(RKDATA) = 4 Then
            Call WriteLog(False, "ROM PACK $05 READ Check")
            cSP.ICMD = 6
            SendParam(1) = "20000B"
            SendParam(2) = "1"
            cSP.IPARA = 2
            cSP.SendPARA = SendParam

            Call cSP.TestComSendRecv()
            Call WriteLog(False, "Send:" & cSP.SEND_TESTCOMMAND)
            Call WriteLog(False, "Recv:" & cSP.RECEIVE_DATA)
            RecvDATA = Split(cSP.RECEIVE_DATA, ",")
            Call WriteLog(False, "ROM PACK $05 DATA = " & RecvDATA(5))
            If RecvDATA(5) <> RKCHK4 Then
                mErrorComment = "ROM PACK $05 READ Check ERROR"
                cSP = Nothing                             'V3.02
                Call WriteLog(False, mErrorComment)
                Throw New ApplicationException(mErrorComment)
            End If
            Call WriteLog(False, "ROM PACK $05 READ Check OK!")

        End If

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
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\RK\RK_SetUp.ppt"
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
            dfrm.DISP_FILENAME = mAppPath & "900 Picture\ppt\RK\RK_Fin.ppt"
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

    Public Sub LabelPrint(ByVal dbName As String, ByVal rDTkey As String)

        Dim wTime As Double = 0

        Dim tbName As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim rField As String
        'Dim kField As String

        Dim sModel As String
        Dim sXJNO As String
        Dim sHRev As String
        Dim sSRev As String
        Dim sLSer As String
        Dim sDate As String
        Dim sInsModel As String
        Dim sEquipmentNo As String

        Dim rDate As Date
        Dim nwYear As String
        Dim nwMonth As String
        Dim nwDay As String

        Dim mPrinterName As String
        Dim mPrintModel, mPrintSUFFIX, mPrintSupply, mPrintInput1, mPrintInput2, mPrintOutput1, mPrintOutput2 As String
        Dim mPrintUL, mPrintCE, mPrintWEEE, mPrintCTick, mPrintRoHS, mPrintModuleName As String
        Dim mRModel As String
        Dim mPrintKCNo As String

        '==========================================================================================

        Try
            '印字データの読込み -----------------------------
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName
            'rField = "Mscode,XJNo,ProductHardRev,ProductSoftRev,LineSerial,StartDateTime,InspectionModel"
            'rField = "Mscode,XJNo,ProductHardRev,ProductSoftRev,SerialNo,StartDateTime,InspectionModel"
            rField = "Mscode,XJNo,ProductHardRev,ProductSoftRev,SerialNo,StartDateTime,InspectionModel,EquipmentNo"

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE "
            sSQL &= "StartDateTime LIKE '" & rDTkey & "'"
            sSQL &= " ORDER BY StartDateTime DESC;"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT = 0 Then
                mErrorComment = "主銘板に印字するデータを取得できませんでした。"
                Throw New ApplicationException(mErrorComment)
            ElseIf sData(0) = "" Or sData(2) = "" Or sData(3) = "" Or sData(4) = "" Or sData(5) = "" Then
                mErrorComment = "主銘板に印字するデータに不備があります。"
                Throw New ApplicationException(mErrorComment)
            End If

            sModel = sData(0) : sXJNO = sData(1) : sHRev = sData(2) : sSRev = sData(3) : sLSer = sData(4) : sInsModel = sData(6) : sEquipmentNo = sData(7)

            rDate = DateValue(sData(5))
            sDate = rDate.ToString("yyyy/MM/dd")

            nwYear = Strings.Right(rDate.ToString("yyyy"), 1)
            nwMonth = Mid("ABCDEFGHJKLM", CInt(rDate.Month.ToString), 1)
            nwDay = rDate.ToString("dd")

            cADO = Nothing

            'MSCODE.mdbファイルの読み込み----------------------------------------
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = mAppPath & "400 Common File\MSCODE.mdb"
            tbName = "tblMSCODE"
            cADO.TB_NAME = tbName
            rField = "ModuleName,Model6,Suffix,Supply,Input1,Input2,Output1,Output2,RoHSMark,ULMark,CEMark,WEEEMark,CTickMark,Model,KCNO"        '
            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model = '" & sInsModel & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)
            cADO = Nothing

            mPrintModuleName = sData(0) : mPrintModel = sData(1) : mPrintSUFFIX = sData(2)
            mPrintSupply = sData(3) : mPrintInput1 = sData(4) : mPrintInput2 = sData(5) : mPrintOutput1 = sData(6) : mPrintOutput2 = sData(7)
            mPrintRoHS = sData(8) : mPrintUL = sData(9) : mPrintCE = sData(10) : mPrintWEEE = sData(11) : mPrintCTick = sData(12)
            mRModel = sData(13) : mPrintKCNo = sData(14)

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

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            cADO = Nothing

        End Try

        Try
            cLbl = New cls011_Excel_Operation
            With cLbl

                .FILE_PATH = mAppPath & "410 Print Form\" & "ROMPACKLabel for KC.xls"    '"FAM3-STDLabel.xls"  
                .SHEET_NAME = "Sheet1"
                .EXCEL_START()

                .CELL_DATA_SET("MODULENAME", mPrintModuleName)
                .CELL_DATA_SET("MODEL", mPrintModel)
                .CELL_DATA_SET("SUFFIX1", mPrintSUFFIX)

                .CELL_DATA_SET("REV", sSRev & "：" & sHRev)

                .CELL_DATA_SET("DATE", sDate)
                .CELL_DATA_SET("FAM3SERIAL", sLSer)

                If InStr(sModel, "/CT") <> 0 Then
                    .CELL_DATA_SET("FAM3SERIALH", sLSer & " CT")
                    .CELL_DATA_SET("CT", "CT")
                Else
                    .CELL_DATA_SET("FAM3SERIALH", sLSer)
                End If

                '.CELL_DATA_SET("BARCODE", Mid(sLSer, 2, 8) & Mid(sModel, 1, InStr(sModel, "-") - 1) & Mid(sModel, InStr(sModel, "-") + 1, 5))
                '.CELL_DATA_SET("BARCODE", Mid(sLSer, 2, 8) & Strings.Left(mRModel, InStr(mRModel, "-") - 1) & Mid(mRModel, InStr(mRModel, "-") + 1, mRModel.Length - InStr(mRModel, "-")))        '& mPrintModel & mPrintSUFFIX
                '.CELL_DATA_SET("BARCODE", sLSer & sEquipmentNo & nwYear & nwMonth & nwDay & Strings.Left(mRModel, InStr(mRModel, "-") - 1) & Mid(mRModel, InStr(mRModel, "-") + 1, mRModel.Length - InStr(mRModel, "-")))   ' & Mid(mRModel, 1, 6) & Mid(mRModel, 8, mRModel.Length - 7))       'v3.09でdel
                '.BARCODE_SHOW("BARCODE", 120 + 190 - 120, 260, 120, 27)
                '.BARCODE_SHOW("BARCODE", 120 + 190 - 210, 180, 150, 27)    'v3.09でdel
                Dim bc1 As DotNetBarcode = New DotNetBarcode
                bc1.Type = DotNetBarcode.Types.QRCode
                bc1.QRCopyToClipboard(sLSer & sEquipmentNo & nwYear & nwMonth & nwDay & Strings.Left(mRModel, InStr(mRModel, "-") - 1) & Mid(mRModel, InStr(mRModel, "-") + 1, mRModel.Length - InStr(mRModel, "-")), 7)

                .PICTURE_QRPASTE_CONTROL()
                .PICTURE_QRPASTE_CONTROL2()   'Rev0.14ADD

                If mPrintRoHS = "True" Then .CELL_DATA_SET("ROHS", "'RoHS") '.CELL_DATA_SET("ROHS", "'RoHS")
                If mPrintCE = "False" Then .PICTURE_VISIBLE_CONTROL("CE", False)
                If mPrintCTick = "False" Then .PICTURE_VISIBLE_CONTROL("C-Tick", False)
                If mPrintWEEE = "False" Then .PICTURE_VISIBLE_CONTROL("WEEE", False)

                If mPrintKCNo <> "" Then
                    .CELL_DATA_SET("KCNO", mPrintKCNo)
                Else
                    .PICTURE_VISIBLE_CONTROL("KC", False)
                End If

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
