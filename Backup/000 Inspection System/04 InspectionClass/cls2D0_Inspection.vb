Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls2D0_Inspection
'Function   : Adjustment and Inspection ATE D
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'           :
'Issue      : 2008/12/30  Rev.0    T.Yanagawa       New Release
'================================================================================================================================

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application
'Shared InspectionForm
Imports Inspection_System.frm020_InspectionStartEnd

Public Class cls2D0_Inspection

    'Class ================================================
    Dim cInsF3SP6x As cls2D1_F3SP6x_Inspection
    Dim cInsF3SP7x As cls2D2_F3SP7x_Inspection
    Dim cInsF3RPxx As cls2D3_F3RPxx_Inspection
    Dim cInsF3SPxxPF As cls2D4_F3SPxx_PF_Inspection
    Dim cInsRK As cls2D5_RK_RomPack_Inspection
    Dim cInsF3SPxxKP As cls2D6_F3SPxx_KP_Inspection
    Dim cInsF3BPxx As cls2D7_F3BPxx_Inspcetion
    Dim cInsF3LE1x As cls2D8_F3LE1x_inspection

    Dim cLnk As cls080_F3LinkToLE
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl

    'Local ================================================
    Dim mErrorComment As String
    Dim mErrorCode As String
    Dim mInspectionNo As String
    Dim mInspectionMODEL As String
    Dim mTestSpec As String 'Rev0.20Add
    Dim mTestMeasure As String 'Rev0.20Add

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

    Public ReadOnly Property TestSpec() As String   'Rev0.20Add
        Get
            TestSpec = mTestSpec
        End Get
    End Property

    Public ReadOnly Property TestMeasure() As String   'Rev0.20Add
        Get
            TestMeasure = mTestMeasure
        End Get
    End Property

    Public WriteOnly Property InspectionMODEL() As String
        Set(ByVal Value As String)
            mInspectionMODEL = Value
        End Set
    End Property

    Public WriteOnly Property InspectionNumber() As String
        Set(ByVal Value As String)
            mInspectionNo = Value
        End Set
    End Property

    Public WriteOnly Property InspectionName() As String
        Set(ByVal Value As String)
            iName = Value
        End Set
    End Property

    Public Sub TestExecute()

        'v0.02で見直し
        If InStr(mInspectionMODEL, "F3SP") <> 0 Then
            Select Case Mid(mInspectionMODEL, 5, 1)
                Case "6"
                    'F3SP6x検査
                    Try
                        cInsF3SP6x = New cls2D1_F3SP6x_Inspection
                        Call cInsF3SP6x.TestExecute(iName)
                    Catch ex As Exception
                        If cInsF3SP6x.ERR_COMMENT <> "" Then
                            If InStr(cInsF3SP6x.ERR_COMMENT, ",") <> 0 Then
                                mErrorComment = Replace(cInsF3SP6x.ERR_COMMENT, ",", ".")
                            Else
                                mErrorComment = cInsF3SP6x.ERR_COMMENT
                            End If
                        End If
                        Throw New ApplicationException(mErrorComment)
                    Finally
                        cInsF3SP6x = Nothing
                    End Try
                Case "7"
                    'F3SP7x検査
                    Try
                        cInsF3SP7x = New cls2D2_F3SP7x_Inspection
                        Call cInsF3SP7x.TestExecute(iName)
                    Catch ex As Exception
                        If cInsF3SP7x.ERR_COMMENT <> "" Then
                            If InStr(cInsF3SP7x.ERR_COMMENT, ",") <> 0 Then
                                mErrorComment = Replace(cInsF3SP7x.ERR_COMMENT, ",", ".")
                            Else
                                mErrorComment = cInsF3SP7x.ERR_COMMENT
                            End If
                            mTestSpec = cInsF3SP7x.TestSpec                                         'Rev0.20add
                            mTestMeasure = cInsF3SP7x.TestMeasure                                   'Rev0.20add
                        End If
                        Throw New ApplicationException(mErrorComment)
                    Finally
                        cInsF3SP7x = Nothing
                    End Try
                Case Else
                    Select Case Mid(mInspectionMODEL, 3, 4)
                        Case "SP05", "SP08", "SP21", "SP25", "SP27", "SP35", "SP37"
                            'F3SPxx-PF検査
                            Try
                                cInsF3SPxxPF = New cls2D4_F3SPxx_PF_Inspection
                                Call cInsF3SPxxPF.TestExecute(iName)
                            Catch ex As Exception
                                If cInsF3SPxxPF.ERR_COMMENT <> "" Then
                                    If InStr(cInsF3SPxxPF.ERR_COMMENT, ",") <> 0 Then
                                        mErrorComment = Replace(cInsF3SPxxPF.ERR_COMMENT, ",", ".")
                                    Else
                                        mErrorComment = cInsF3SPxxPF.ERR_COMMENT
                                    End If
                                    mTestSpec = cInsF3SPxxPF.TestSpec                                         'Rev0.20add
                                    mTestMeasure = cInsF3SPxxPF.TestMeasure                                   'Rev0.20add
                                End If
                                Throw New ApplicationException(mErrorComment)
                            Finally
                                cInsF3SPxxPF = Nothing
                            End Try
                            '-----< Rev0.12AD >-----
                        Case "SPV9"
                            'F3SPV9検査
                            Try
                                cInsF3SP7x = New cls2D2_F3SP7x_Inspection
                                Call cInsF3SP7x.TestExecute(iName)
                            Catch ex As Exception
                                If cInsF3SP7x.ERR_COMMENT <> "" Then
                                    If InStr(cInsF3SP7x.ERR_COMMENT, ",") <> 0 Then
                                        mErrorComment = Replace(cInsF3SP7x.ERR_COMMENT, ",", ".")
                                    Else
                                        mErrorComment = cInsF3SP7x.ERR_COMMENT
                                    End If
                                    mTestSpec = cInsF3SP7x.TestSpec                                         'Rev0.20add
                                    mTestMeasure = cInsF3SP7x.TestMeasure                                   'Rev0.20add
                                End If
                                Throw New ApplicationException(mErrorComment)
                            Finally
                                cInsF3SP7x = Nothing
                            End Try
                            '------------------
                        Case Else
                            'F3SPxx-KP検査
                            Try
                                cInsF3SPxxKP = New cls2D6_F3SPxx_KP_Inspection
                                Call cInsF3SPxxKP.TestExecute(iName)
                            Catch ex As Exception
                                If cInsF3SPxxKP.ERR_COMMENT <> "" Then
                                    If InStr(cInsF3SPxxKP.ERR_COMMENT, ",") <> 0 Then
                                        mErrorComment = Replace(cInsF3SPxxKP.ERR_COMMENT, ",", ".")
                                    Else
                                        mErrorComment = cInsF3SPxxKP.ERR_COMMENT
                                    End If
                                    mTestSpec = cInsF3SPxxKP.TestSpec                                         'Rev0.20add
                                    mTestMeasure = cInsF3SPxxKP.TestMeasure                                   'Rev0.20add
                                End If
                                Throw New ApplicationException(mErrorComment)
                            Finally
                                cInsF3SPxxKP = Nothing
                            End Try

                    End Select

            End Select

        ElseIf InStr(mInspectionMODEL, "F3RP") <> 0 Then
            'F3RPxx検査
            Try
                cInsF3RPxx = New cls2D3_F3RPxx_Inspection
                Call cInsF3RPxx.TestExecute(iName)
            Catch ex As Exception
                If cInsF3RPxx.ERR_COMMENT <> "" Then
                    If InStr(cInsF3RPxx.ERR_COMMENT, ",") <> 0 Then
                        mErrorComment = Replace(cInsF3RPxx.ERR_COMMENT, ",", ".")
                    Else
                        mErrorComment = cInsF3RPxx.ERR_COMMENT
                    End If
                End If
                Throw New ApplicationException(mErrorComment)
            Finally
                cInsF3RPxx = Nothing
            End Try

        ElseIf Strings.Left(mInspectionMODEL, 2) = "RK" Then
            'ﾛﾑﾊﾟｯｸ検査
            Try
                cInsRK = New cls2D5_RK_RomPack_Inspection
                Call cInsRK.TestExecute(iName)
            Catch ex As Exception
                If cInsRK.ERR_COMMENT <> "" Then
                    If InStr(cInsRK.ERR_COMMENT, ",") <> 0 Then
                        mErrorComment = Replace(cInsRK.ERR_COMMENT, ",", ".")
                    Else
                        mErrorComment = cInsRK.ERR_COMMENT
                    End If
                End If
                Throw New ApplicationException(mErrorComment)
            Finally
                cInsRK = Nothing
            End Try

        ElseIf InStr(mInspectionMODEL, "F3BP") <> 0 Then
            'F3BPxx検査
            Try
                cInsF3BPxx = New cls2D7_F3BPxx_Inspcetion
                Call cInsF3BPxx.TestExecute(iName)
            Catch ex As Exception
                If cInsF3BPxx.ERR_COMMENT <> "" Then
                    If InStr(cInsF3BPxx.ERR_COMMENT, ",") <> 0 Then
                        mErrorComment = Replace(cInsF3BPxx.ERR_COMMENT, ",", ".")
                    Else
                        mErrorComment = cInsF3BPxx.ERR_COMMENT
                    End If
                    mTestSpec = cInsF3BPxx.TestSpec                                         'Rev0.20add
                    mTestMeasure = cInsF3BPxx.TestMeasure                                   'Rev0.20add
                End If
                Throw New ApplicationException(mErrorComment)
            Finally
                cInsF3BPxx = Nothing
            End Try

        ElseIf InStr(mInspectionMODEL, "F3LE") <> 0 Then
            'F3LE1x検査
            Try
                cInsF3LE1x = New cls2D8_F3LE1x_inspection
                Call cInsF3LE1x.TestExecute(iName)
            Catch ex As Exception
                If cInsF3LE1x.ERR_COMMENT <> "" Then
                    If InStr(cInsF3LE1x.ERR_COMMENT, ",") <> 0 Then
                        mErrorComment = Replace(cInsF3LE1x.ERR_COMMENT, ",", ".")
                    Else
                        mErrorComment = cInsF3LE1x.ERR_COMMENT
                    End If
                End If
                Throw New ApplicationException(mErrorComment)
            Finally
                cInsF3LE1x = Nothing
            End Try

        Else
            mErrorComment = "形名のプログラムが存在しません。確認してください。"
            MsgBox(mErrorComment)
            Throw New ApplicationException(mErrorComment)

        End If

    End Sub

    Public Sub PowerOFF() '*******異常があれば電源OFF、FA-M3で電源制御できるもの*******

        Try
            cBUS = New cls122_FAM3_IOBUS_Ctrl
            Call cBUS.BUSMONITOR_TB04_END()
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Throw New ApplicationException(mErrorComment)
        Finally
            cBUS = Nothing
        End Try

    End Sub

    Public Sub LabelPrint(ByVal dbname As String, ByVal dtkey As String)

        'v0.02で見直し
        If InStr(mInspectionMODEL, "F3SP") <> 0 Then
            Select Case Mid(mInspectionMODEL, 5, 1)
                Case "6"
                    'F3SP6x
                    Try
                        cInsF3SP6x = New cls2D1_F3SP6x_Inspection
                        Call cInsF3SP6x.LabelPrint(dbname, dtkey)
                    Catch ex As Exception
                    Finally
                        cInsF3SP6x = Nothing
                    End Try
                Case "7"
                    'F3SP7x
                    Try
                        cInsF3SP7x = New cls2D2_F3SP7x_Inspection
                        Call cInsF3SP7x.LabelPrint(dbname, dtkey)
                    Catch ex As Exception
                    Finally
                        cInsF3SP7x = Nothing
                    End Try
                Case Else
                    Select Case Mid(mInspectionMODEL, 3, 4)
                        Case "SP05", "SP08", "SP21", "SP25", "SP27", "SP35", "SP37"
                            'F3SPxx-PF検査
                            Try
                                cInsF3SPxxPF = New cls2D4_F3SPxx_PF_Inspection
                                Call cInsF3SPxxPF.LabelPrint(dbname, dtkey)
                            Catch ex As Exception
                            Finally
                                cInsF3SPxxPF = Nothing
                            End Try
                            '-----<Rev0.12AD>-----
                        Case "SPV9"
                            'F3SPV9
                            Try
                                cInsF3SP7x = New cls2D2_F3SP7x_Inspection
                                Call cInsF3SP7x.LabelPrint(dbname, dtkey)
                            Catch ex As Exception
                            Finally
                                cInsF3SP7x = Nothing
                            End Try
                            '---------------------
                        Case Else
                            'F3SPxx-KP
                            Try
                                cInsF3SPxxKP = New cls2D6_F3SPxx_KP_Inspection
                                Call cInsF3SPxxKP.LabelPrint(dbname, dtkey)
                            Catch ex As Exception
                            Finally
                                cInsF3SPxxKP = Nothing
                            End Try
                    End Select

            End Select

        ElseIf InStr(mInspectionMODEL, "F3RP") <> 0 Or Strings.Left(mInspectionMODEL, 4) = "F3LE" Then
            'F3RPxx
            Try
                cInsF3RPxx = New cls2D3_F3RPxx_Inspection
                Call cInsF3RPxx.LabelPrint(dbname, dtkey)
            Catch ex As Exception
            Finally
                cInsF3RPxx = Nothing
            End Try

        ElseIf Strings.Left(mInspectionMODEL, 2) = "RK" Then
            'ﾛﾑﾊﾟｯｸ
            Try
                cInsRK = New cls2D5_RK_RomPack_Inspection
                Call cInsRK.LabelPrint(dbname, dtkey)
            Catch ex As Exception
            Finally
                cInsRK = Nothing
            End Try

            '-----< Rev0.10ADD >-----
        ElseIf InStr(mInspectionMODEL, "F3BP") <> 0 Then
            'F3BPxx
            Try
                cInsF3SPxxPF = New cls2D4_F3SPxx_PF_Inspection
                Call cInsF3SPxxPF.LabelPrint(dbname, dtkey)
            Catch ex As Exception
            Finally
                cInsF3SPxxPF = Nothing
            End Try

            '------------------------

        Else
            mErrorComment = "形名のプログラムが存在しません。確認してください。"
            MsgBox(mErrorComment)
            Throw New ApplicationException(mErrorComment)

        End If

    End Sub

    Public Sub BUS_MONITOR_POWER_OFF()

        Try
            cBUS = New cls122_FAM3_IOBUS_Ctrl
            Call cBUS.BUS_MONITOR_POWER_OFF()
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Throw New ApplicationException(mErrorComment)
        Finally
            cBUS = Nothing
        End Try

    End Sub

    Public Sub BUS_MONITOR_POWER_ON()

        Try
            cBUS = New cls122_FAM3_IOBUS_Ctrl
            Call cBUS.BUS_MONITOR_POWER_ON()
        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            Throw New ApplicationException(mErrorComment)
        Finally
            cBUS = Nothing
        End Try

    End Sub

End Class
