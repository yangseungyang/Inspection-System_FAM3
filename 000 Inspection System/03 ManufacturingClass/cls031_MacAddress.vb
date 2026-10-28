Option Strict On : Option Explicit On

Public Class cls031_MacAddress

    Dim mMacAdd As String
    Dim mError As String
    Dim mDBName As String
    Dim mSerial As String

    'Address ===============================
    Public ReadOnly Property MacAddress() As String
        Get
            MacAddress = mMacAdd
        End Get
    End Property

    'Get Error =============================
    Public ReadOnly Property ERR_MSG() As String
        Get
            ERR_MSG = mError
        End Get
    End Property

    'シリアル番号
    Public Property Serial() As String
        Get
            Serial = mSerial
        End Get
        Set(ByVal value As String)
            mSerial = value
        End Set
    End Property

    'MACアドレス用データベースファイル名
    Public Property DBFileName() As String
        Get
            DBFileName = mDBName
        End Get
        Set(ByVal value As String)
            mDBName = value
        End Set
    End Property

    Public Sub IssueMacAddress()
        '==========================================================================================
        ' F.Name   : IssueMacAddress
        ' Function : Issue MacAddress
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim cSQL As cls010_MDB_Operation

        Dim dbName As String
        Dim tbName As String
        Dim sSQL As String
        Dim rField As String
        Dim kField As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim dbFixed As String
        Dim dbNumber As String
        Dim dbHiLim As String
        Dim dbLoLim As String
        'Dim nwNumber As String
        Dim vNextNum As Long
        Dim vHiLim As Long
        Dim vLoLim As Long

        'Read DB MacAddress ========================================
        cSQL = New cls010_MDB_Operation

        dbName = mDBName
        tbName = "tblNextMAC"
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName

        cSQL.DB_PATH = dbName
        cSQL.TB_NAME = tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        dbFixed = sData(0) : dbNumber = sData(1) : dbHiLim = sData(2) : dbLoLim = sData(3)

        vNextNum = Convert.ToInt32(dbNumber, 16)        '"&H" & dbNumber
        vHiLim = Convert.ToInt32(dbHiLim, 16)           '"&H" & dbHiLim
        vLoLim = Convert.ToInt32(dbLoLim, 16)           '"&H" & dbLoLim

        'Hi/Lo Check ---------------------------------
        If vNextNum > vHiLim Then
            mError = "This Mac Address exceeded the upper bound.!!!" & "上限を超えています。発番できません。"
            GoTo ErrorTrap
        End If

        If vNextNum < vLoLim Then
            mError = "This Mac Address exceeded the lower bound.!!!" & "下限を超えています。発番できません。"
            GoTo ErrorTrap
        End If

        'Fixed number check =========================================
        If dbFixed <> "000064" Then
            mError = "MAC Address Fixed number NG : " & dbFixed
            GoTo ErrorTrap
        End If

        'This time Mac Address ======================================
        mMacAdd = dbFixed & dbNumber

        'Search past Mac Address ====================================
        cSQL = New cls010_MDB_Operation

        dbName = mDBName
        tbName = "tblPastMAC"
        rField = "*" : kField = "PastMac"

        sSQL = "SELECT " & "*" & " FROM " & tbName & " "
        sSQL = sSQL & "WHERE " & kField & " LIKE " & "'" & mMacAdd & "' "
        sSQL = sSQL & "ORDER BY " & kField

        cSQL.DB_PATH = dbName
        cSQL.TB_NAME = tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.DB_DATA_COUNT <> 0 Then
            mError = "This MAC Address has been issued in the past." & "発行済番号です。発番できません。"
            GoTo ErrorTrap
        End If

        cSQL = Nothing

        Exit Sub

ErrorTrap:

        If mError = "" Then mError = "Issue Mac Address NG"

    End Sub

    Public Sub IssueNextMacAddress()
        '==========================================================================================
        ' F.Name   : IssueNextMacAddress
        ' Function : Issue next MacAddress
        ' Data     : Nothing
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim cSQL As cls010_MDB_Operation

        Dim i As Integer
        Dim dbName As String
        Dim tbName As String
        Dim sSQL As String
        Dim rField As String
        Dim kField As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim dbFixed As String
        Dim dbNumber As String
        Dim dbHiLim As String
        Dim dbLoLim As String
        Dim nwNumber As String
        Dim vNextNum As Long
        Dim vHiLim As Long
        Dim vLoLim As Long

        '---- Ver 0.21 JDI ADD ----
        Dim First_ALM As Integer
        Dim Last_ALM As Integer
        Dim ALM_Count As Integer
        '--------------------------

        'Read DB MacAddress ========================================
        cSQL = New cls010_MDB_Operation

        dbName = mDBName
        tbName = "tblNextMAC"
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName

        cSQL.DB_PATH = dbName
        cSQL.TB_NAME = tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        dbFixed = sData(0) : dbNumber = sData(1) : dbHiLim = sData(2) : dbLoLim = sData(3)

        vNextNum = Convert.ToInt32(dbNumber, 16)        '"&H" & dbNumber
        vHiLim = Convert.ToInt32(dbHiLim, 16)           '"&H" & dbHiLim
        vLoLim = Convert.ToInt32(dbLoLim, 16)           '"&H" & dbLoLim

        'Past Mac Address ======================================
        mMacAdd = dbFixed & dbNumber

        'Issue Next Mac Address =====================================
        vNextNum = vNextNum + 1         '("&H" & dbNumber) + 1

        '------------------------- Ver 0.21 JDI ADD Start --------------------------
        cSQL = New cls010_MDB_Operation

        dbName = mDBName
        tbName = "tblMacAlarm"
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName

        cSQL.DB_PATH = dbName
        cSQL.TB_NAME = tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        First_ALM = Convert.ToInt32(sData(0), 10)
        Last_ALM = Convert.ToInt32(sData(1), 10)
        ALM_Count = Convert.ToInt32(sData(2), 10)

        Try
            If vNextNum > vHiLim Then
                MsgBox("Mac 주소의 다음 번호가 상한치를 넘어갔습니다." & vbCrLf & " 라인리더를 불러주세요.")
                Throw New ApplicationException("The next number exceeded the upper bound.!!!" & "次回の番号が上限を超えました。")
            End If

            If vNextNum < vLoLim Then
                MsgBox("Mac 주소의 다음 번호가 하안치를 넘어갔습니다." & vbCrLf & " 라인리더를 불러주세요.")
                Throw New ApplicationException("The next number exceeded the upper bound.!!!" & "次回の番号が下限を超えました。")
            End If

            If vNextNum > vHiLim - First_ALM And vNextNum < vHiLim And ALM_Count = 1 Then
                MsgBox("Mac 주소가 " & (vHiLim - vNextNum).ToString & "개 남아있습니다." & vbCrLf _
                       & "라인리더를 불러 주세요.", MsgBoxStyle.Critical)
            ElseIf vNextNum > vHiLim - Last_ALM And vNextNum < vHiLim Then
                MsgBox("Mac 주소가 " & (vHiLim - vNextNum).ToString & "개 남아있습니다." & vbCrLf _
                       & "라인리더를 불러 주세요.", MsgBoxStyle.Critical)
            End If

        Catch ex As Exception

        Finally

        End Try
        '-------------------------- Ver 0.21 JDI ADD End ---------------------------

        '------------------------- Ver 0.21 JDI DEL Start --------------------------
        'If vNextNum > vHiLim Then
        '    MsgBox("The next number exceeded the upper bound.!!!" & "次回の番号が上限を超えました。")
        'End If

        'If vNextNum < vLoLim Then
        '    MsgBox("The next number exceeded the upper bound.!!!" & "次回の番号が下限を超えました。")
        'End If

        'If vNextNum > vHiLim - 1000 And vNextNum < vHiLim Then
        '    MsgBox("The number has decreased.!!!" & "MACｱﾄﾞﾚｽが残り少数です。" & vbCrLf & "1度必ず生産技に連絡してください！" _
        '            & "残り" & (vHiLim - vNextNum).ToString & "台です。", MsgBoxStyle.Critical)
        'End If

        '-------------------------- Ver 0.21 JDI DEL End ---------------------------

        nwNumber = Hex(vNextNum)
        For i = 1 To 6 - Len(nwNumber)
            nwNumber = 0 & nwNumber
        Next


        'Addition past Mac Address =================================
        cSQL = New cls010_MDB_Operation

        dbName = mDBName
        tbName = "tblPastMAC"
        rField = "*" : kField = mMacAdd

        sSQL = "INSERT INTO " & tbName
        If mSerial = "" Then
            sSQL = sSQL & " ( PastMAC ) VALUES ('" & mMacAdd & "');"
        Else
            sSQL = sSQL & " ( PastMAC , Serial ) VALUES ('" & mMacAdd & "','" & mSerial & "');"
        End If
        cSQL.DB_PATH = dbName
        cSQL.TB_NAME = tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_EXEC_JET()
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing


        'Rewrite Next Mac Address ===================================
        cSQL = New cls010_MDB_Operation

        dbName = mDBName
        tbName = "tblNextMAC"
        rField = "*" : kField = ""

        sSQL = "SELECT " & rField & " FROM " & tbName

        cSQL.DB_PATH = dbName
        cSQL.TB_NAME = tbName
        cSQL.DB_SQL = sSQL

        Call cSQL.DB_READ_JET(fName, sData)
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        sSQL = "UPDATE " & tbName & " SET "
        sSQL = sSQL & fName(0) & " = " & "'" & dbFixed & "' , "
        sSQL = sSQL & fName(1) & " = " & "'" & nwNumber & "'"
        sSQL = sSQL & " WHERE "
        sSQL = sSQL & fName(0) & " = " & "'" & sData(0) & "'"

        cSQL.DB_SQL = sSQL

        Call cSQL.DB_EXEC_JET()
        If cSQL.ERR_COMMENT <> "" Then GoTo ErrorTrap

        cSQL = Nothing

        Exit Sub

ErrorTrap:

        If mError = "" Then mError = "Issue next Mac Address NG"

    End Sub

End Class
