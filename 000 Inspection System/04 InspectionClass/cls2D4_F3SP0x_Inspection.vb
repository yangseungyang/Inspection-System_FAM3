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

Public Class cls2D4_F3SP0x_Inspection
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

    Dim dfrm As frm500_MessagePopUp

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String

    Dim rs232cSPPort As String = "COM4"
    'Dim rs232cCPU1SPPort As String = "COM4"

    Dim IP_U1 As String = "192.168.1.2"         '制御用FA-M3のLE
    Dim IP_U2 As String = "192.168.1.11"        'ﾒｲﾝCPUのLE
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

            Case "I50" : Call I50(iName)            'I50:SLOT/XDFAIL TEST
            Case "I51" : Call I51(iName)            'I51:ACS SIGNAL TEST
            Case "I52" : Call I52(iName)            'I52:BCLR TEST
            Case "I53" : Call I53(iName)            'I53:BR TEST
                'Case "I54" : Call I54(iName)            'I54:I/O BUS WRITE/READ TEST
            Case "I55" : Call I55(iName)            'I55:I/O BUS INTERRUPT TEST
            Case "I57" : Call I57(iName)            'I57:XPFAIL TEST
            Case "I58" : Call I58(iName)            'I58:XSFAIL TEST
            Case "I59" : Call I59(iName)            'I59:IOﾊﾞｽ検査前準備


            Case "I90" : Call I90(iName)            'I90:ﾕﾆｯﾄ電源ON
            Case "I91" : Call I91(iName)            'I91:ﾕﾆｯﾄ電源OFF
                'Case "I92" : Call I92(iName)            'I92:検査開始時処理
                'Case "I93" : Call I93(iName)            'I93:検査終了時メッセージ

            Case Else : mErrorComment = "検査プロシージャが定義されていません。"
                MsgBox(mErrorComment)

        End Select

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

        Call WriteLog(False, "ｽﾛｯﾄｱﾄﾞﾚｽ & XDFAILﾃｽﾄ中・・・約50秒")
        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = 5
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(60)
            If cPF.SlotAddressTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "ｽﾛｯﾄｱﾄﾞﾚｽﾃｽﾄ異常:" & cPF.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
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
        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = 5
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cPF.XACSTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XACSﾃｽﾄ異常:" & cPF.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
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
        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = 5
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cPF.XBCLRTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XBCLRﾃｽﾄ異常:" & cPF.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
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
        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = 5
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cPF.XBRTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XBRﾃｽﾄ異常:" & cPF.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
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

        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = 5
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cPF.XIRQTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XIRQﾃｽﾄ異常:" & cPF.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
        End Try

        Call WriteLog(False, "XIRQ検査OK")

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

        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = 5
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cPF.XPFAILTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XPFAILﾃｽﾄ異常:" & cPF.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
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

        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = 5
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(20)
            If cPF.XSFAILTest() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "XSFAILﾃｽﾄ異常:" & cPF.ERR_MSG
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
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

        cPF = New cls125_PF_IOBUS_TEST
        cPF.TESTSlotAddress = slot
        cPF.IPAdd = IP_U2

        Try
            Call CountStart(15)
            If cPF.TestPreparation() = False Then
                Throw New ApplicationException(mErrorComment)
            End If
        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽﾃｽﾄ前準備失敗"
            Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call CountStop()
            Call WriteLog(False, cPF.LOG)
            cPF = Nothing
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
End Class
