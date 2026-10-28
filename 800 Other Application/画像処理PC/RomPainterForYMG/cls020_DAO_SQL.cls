VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls020_DAO_SQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Project setting
'Microsoft DAO Object Libraly

'DAO ---------------------------
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim fd As DAO.Field

'Property Local Parameter ------
Dim mFieldName   As String
Dim mFieldCount  As Long
Dim mDBData      As String
Dim mDBDataCount As Long
Dim mRecCount    As Long
Dim mError       As String

'mdb Field Name =====================
Public Property Get DB_FieldName() As String
    DB_FieldName = mFieldName
End Property

'mdb Field Count =====================
Public Property Get DB_FieldCount() As Long
    DB_FieldCount = mFieldCount
End Property

'mdb Field Data ====================
Public Property Get DB_Data() As String
    DB_Data = mDBData
End Property

'mdb Data Count ====================
Public Property Get DB_DataCount() As Long
    DB_DataCount = mDBDataCount
End Property

'mdb Record Count ====================
Public Property Get DB_RecCount() As Long
    DB_RecCount = mRecCount
End Property

'mdb Comm Error ====================
Public Property Get ERR_MSG() As String
    ERR_MSG = mError
End Property

Public Function SQL_EXEC(dbPath As String, dbSQL As String) As String
'==========================================================================================
' F.Name   : SQL_Exec
' Function : SQL Exec
' Data     : DB path , SQL
' Return   : Nothing
' Remarks  :
'==========================================================================================

    Dim Tout  As Double
    Dim Tcnt  As Double
    Dim erNum As Long
        
    mError = ""

    'open DB ============================================
    Tcnt = 0.1
    Do
        Debug.Print ""
        Debug.Print "SET DB OPEN DATABASE"
        On Error Resume Next
        Set db = OpenDatabase(dbPath)
        erNum = Err
        If erNum = 0 Then Exit Do
        If erNum <> 3050 Then GoTo Errortrap
        
        Call mWait(Tcnt): Tout = Tout + Tcnt: Debug.Print Left(CStr(Tout), 3)
        If Tout >= 5 Then GoTo Errortrap
    Loop
    Debug.Print "DB OPEN OK"

    On Error GoTo Errortrap

    'SQL Exec -----------
    Debug.Print dbSQL
    
    SQL_EXEC = ""
    db.Execute dbSQL
    db.Close: Set db = Nothing
    
    Exit Function

Errortrap:

    If Err = 3050 Then
        mError = "ERROR:DB LOCK"
        Exit Function
    End If

    If Err = 3051 Then
        mError = "ERROR:DB ACCESS"
        Exit Function
    End If

    If Err = 3044 Then
        mError = "ERROR:DB PATH"
        Exit Function
    End If
    
    If Err = 3024 Then
        mError = "ERROR:A0 Nothing"
        Exit Function
    End If
    
    mError = "ERROR:SQL EXEC"
    db.Close: Set db = Nothing
    
End Function

Public Function OPEN_REC_SET(dbPath As String, dbSQL As String, dbFname() As String, dbRdata() As String) As String
'==========================================================================================
' F.Name   : OPEN_REC_SET
' Function : OpenRecordset
' Data     : DB path , SQL
' Return   : Field name, data
' Remarks  :
'==========================================================================================
    
    Dim i As Integer, j As Integer
    
    Dim Tout  As Double
    Dim Tcnt  As Double
    Dim erNum As Long
        
        
    mError = ""
    
    'open DB ============================================
    Tcnt = 0.1
    Do
        Debug.Print ""
        Debug.Print "SET DB OPEN DATABASE"
        On Error Resume Next
        Set db = OpenDatabase(dbPath)
        erNum = Err
        If erNum = 0 Then Exit Do
        If erNum <> 3050 Then GoTo Errortrap
        
        Call mWait(Tcnt): Tout = Tout + Tcnt: Debug.Print Left(CStr(Tout), 3)
        If Tout >= 5 Then GoTo Errortrap
    Loop
    Debug.Print "DB OPEN OK"

    On Error GoTo Errortrap

    'open recordset =====================================
    Debug.Print dbSQL
    
    Set rs = db.OpenRecordset(dbSQL)
    mFieldCount = rs.Fields.Count
    mRecCount = rs.RecordCount
    
    Debug.Print mRecCount & "-" & mFieldCount

    'field count & name & "," connect ===================
    ReDim dbFname(rs.Fields.Count - 1) As String
    
    i = 0: mFieldName = ""
    For Each fd In rs.Fields
        dbFname(i) = fd.Name
        mFieldName = mFieldName & dbFname(i) & ","
        i = i + 1
    Next
    mFieldName = Left(mFieldName, Len(mFieldName) - 1)


    'data ===============================================
    mDBData = ""
    mDBDataCount = rs.RecordCount * rs.Fields.Count
    If mDBDataCount = 0 Then
        mError = "ERROR:DATA 0"
        GoTo Errortrap
    End If
    ReDim dbRdata(mDBDataCount - 1) As String
    
    'Record -------------
    For i = 0 To rs.RecordCount - 1
        
        If i <> 0 Then mDBData = mDBData & ","
        
        'Field ----------
        For j = 0 To rs.Fields.Count - 1
            If j <> rs.Fields.Count - 1 Then
                mDBData = mDBData & rs.Fields(j) & ","
            Else
                mDBData = mDBData & rs.Fields(j)
            End If
        Next
        
        rs.MoveNext
    
    Next
    
    If mDBData <> "" Then
        dbRdata() = Split(mDBData, ",")
    Else
        ReDim dbRdata(0) As String
        dbRdata(0) = ""
    End If
    
    ' ====================================================
    OPEN_REC_SET = mFieldName & "," & mDBData
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
        
    Exit Function

Errortrap:
    
    i = Err
    
    If Err = 3050 Then
        mError = "ERROR:DB LOCK"
        Exit Function
    End If
    
    If Err = 3051 Then
        mError = "ERROR:DB ACCESS"
        Exit Function
    End If
    
    If Err = 3044 Then
        mError = "ERROR:DB PATH"
        Exit Function
    End If

    If Err = 3024 Then
        mError = "ERROR:A0 Nothing"
        Exit Function
    End If
    
    If mError = "" Then mError = "ERROR:SQL RECSET"
    db.Close: Set db = Nothing
    
End Function

Public Function OPEN_REC_SET_QUICK(dbPath As String, dbSQL As String, dbFname() As String, dbRdata() As String)
'==========================================================================================
' F.Name   : OPEN_REC_SET_QUICK
' Function : OpenRecordset(,Data Nothing)
' Data     : DB path , SQL
' Return   : Field name(),data()
' Remarks  :
'==========================================================================================
    
    Dim i As Integer, j As Integer
    Dim Tout   As Double
    Dim Tcnt   As Double
    Dim erNum  As Long
    Dim dCount As Long
    
    On Error GoTo Errortrap
    
    mError = ""

    'open DB ============================================
    Tcnt = 0.1
    Do
        Debug.Print ""
        Debug.Print "SET DB OPEN DATABASE"
        On Error Resume Next
        Set db = OpenDatabase(dbPath)
        erNum = Err
        If erNum = 0 Then Exit Do
        If erNum <> 3050 Then GoTo Errortrap
        
        Call mWait(Tcnt): Tout = Tout + Tcnt: Debug.Print Left(CStr(Tout), 3)
        If Tout >= 5 Then GoTo Errortrap
    Loop
    Debug.Print "DB OPEN OK"

    On Error GoTo Errortrap

    'open recordset =====================================
    Debug.Print dbSQL
    
    Set rs = db.OpenRecordset(dbSQL)
    mFieldCount = rs.Fields.Count
    mRecCount = rs.RecordCount
    
    Debug.Print mRecCount & "-" & mFieldCount

    'field count & name & "," connect ===================
    ReDim dbFname(rs.Fields.Count - 1) As String
    
    i = 0: mFieldName = ""
    For Each fd In rs.Fields
        dbFname(i) = fd.Name
        mFieldName = mFieldName & dbFname(i) & ","
        i = i + 1
    Next
    mFieldName = Left(mFieldName, Len(mFieldName) - 1)


    'data ===============================================
    mDBDataCount = rs.RecordCount * rs.Fields.Count
    If mDBDataCount = 0 Then
        mError = "ERROR:DATA 0"
        GoTo Errortrap
    End If
    ReDim dbRdata(mDBDataCount - 1) As String

    dCount = 0
    'Record -------------
    For i = 0 To rs.RecordCount - 1
        
        'Field ----------
        For j = 0 To rs.Fields.Count - 1
            dbRdata(dCount) = rs.Fields(j)
            dCount = dCount + 1
        Next
        
        rs.MoveNext
    
    Next
    
    ' ====================================================
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
        
    Exit Function

Errortrap:

    If Err = 3050 Then
        mError = "ERROR:DB LOCK"
        Exit Function
    End If

    If Err = 3051 Then
        mError = "ERROR:DB ACCESS"
        Exit Function
    End If
    
    If Err = 3044 Then
        mError = "ERROR:DB PATH"
        Exit Function
    End If

    If Err = 3024 Then
        mError = "ERROR:A0 Nothing"
        Exit Function
    End If
    
    If mError = "" Then mError = "ERROR:SQL RECSET"
    db.Close: Set db = Nothing
    
End Function

Public Function OPEN_REC_SET_XLS(dbPath As String, dbSQL As String, dbFname() As String, dbRdata() As String) As String
'==========================================================================================
' F.Name   : OPEN_REC_SET_XLS
' Function : OpenRecordset
' Data     : DB path , SQL
' Return   : Field name, data
' Remarks  :
'==========================================================================================
    
    Dim i As Integer, j As Integer
    
    Dim Tout  As Double
    Dim Tcnt  As Double
    Dim erNum As Long
        
        
    mError = ""
    
    'open DB ============================================
    Tcnt = 0.1
    Do
        Debug.Print ""
        Debug.Print "SET DB OPEN DATABASE"
        On Error Resume Next
'        Set db = OpenDatabase(dbPath)
        Set db = OpenDatabase(dbPath, False, False, "EXCEL 8.0;HDR=YES;")
        erNum = Err
        If erNum = 0 Then Exit Do
        If erNum <> 3050 Then GoTo Errortrap
        
        Call mWait(Tcnt): Tout = Tout + Tcnt: Debug.Print Left(CStr(Tout), 3)
        If Tout >= 5 Then GoTo Errortrap
    Loop
    Debug.Print "DB OPEN OK"

    On Error GoTo Errortrap

    'open recordset =====================================
    Debug.Print dbSQL
    
    Set rs = db.OpenRecordset(dbSQL)
    mFieldCount = rs.Fields.Count
    mRecCount = rs.RecordCount
    
    Debug.Print mRecCount & "-" & mFieldCount

    'field count & name & "," connect ===================
    ReDim dbFname(rs.Fields.Count - 1) As String
    
    i = 0: mFieldName = ""
    For Each fd In rs.Fields
        dbFname(i) = fd.Name
        mFieldName = mFieldName & dbFname(i) & ",,"
        i = i + 1
    Next
    mFieldName = Left(mFieldName, Len(mFieldName) - 1)


    'data ===============================================
    mDBData = ""
    mDBDataCount = rs.RecordCount * rs.Fields.Count
    If mDBDataCount = 0 Then
        mError = "ERROR:DATA 0"
        GoTo Errortrap
    End If
    ReDim dbRdata(mDBDataCount - 1) As String
    
    'Record -------------
    For i = 0 To rs.RecordCount - 1
        
        If i <> 0 Then mDBData = mDBData & ",,"
        
        'Field ----------
        For j = 0 To rs.Fields.Count - 1
            If j <> rs.Fields.Count - 1 Then
                mDBData = mDBData & rs.Fields(j) & ",,"
            Else
                mDBData = mDBData & rs.Fields(j)
            End If
        Next
        
        rs.MoveNext
    
    Next
    
    If mDBData <> "" Then
        dbRdata() = Split(mDBData, ",,")
    Else
        ReDim dbRdata(0) As String
        dbRdata(0) = ""
    End If
    
    ' ====================================================
    OPEN_REC_SET_XLS = mFieldName & "," & mDBData
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
        
    Exit Function

Errortrap:
    
    i = Err
    
    If Err = 3050 Then
        mError = "ERROR:DB LOCK"
        Exit Function
    End If
    
    If Err = 3051 Then
        mError = "ERROR:DB ACCESS"
        Exit Function
    End If
    
    If Err = 3044 Then
        mError = "ERROR:DB PATH"
        Exit Function
    End If

    If Err = 3024 Then
        mError = "ERROR:A0 Nothing"
        Exit Function
    End If
    
    If mError = "" Then mError = "ERROR:SQL RECSET"
    db.Close: Set db = Nothing
    
End Function


Private Function mWait(Sec As Double)
'==========================================================================================
' F.Name   : mWait
' Function : Sec Timer
' Data     : 0.000 sec
' Return   : Nothing
' Remarks  :
'==========================================================================================

    Dim StartTime As Double
    Dim NowTime   As Double
    Dim ErrTime   As Double
    
    StartTime = Timer
    Do While Sec > ErrTime
        NowTime = Timer
        ErrTime = NowTime - StartTime
        If ErrTime < 0 Then ErrTime = 24# * 3600# - StartTime + NowTime
    Loop
        
End Function

