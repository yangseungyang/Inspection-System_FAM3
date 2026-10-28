Option Strict On : Option Explicit On


'Doevents
Imports System.Windows.Forms.Application

Public Class cls124_KP_IOBUS_TEST

    Dim cMcm As cls120_SP_M_Command
    Dim cBUS As cls122_FAM3_IOBUS_Ctrl

    Dim mSlotAddress As Integer = 0
    Dim mLOG As String = ""
    Dim mErrorComment As String
    Dim mIPAdd As String = ""

    Dim mRegSlotAddress As String = "00000200"
    Dim mRegIRQMon As String = "00000207"
    Dim mRegBRMon As String = "00000230"
    Dim mRegMODFail As String = "00000224"
    Dim mRegLFail As String = "00000217"

    'IPAddress -----------------------------
    Public Property IPAdd() As String
        Set(ByVal Value As String)
            mIPAdd = Value
        End Set
        Get
            IPAdd = mIPAdd
        End Get
    End Property

    '検査ｽﾛｯﾄｱﾄﾞﾚｽ設定 ====================================
    Public WriteOnly Property TESTSlotAddress() As Integer
        Set(ByVal rp As Integer)
            mSlotAddress = rp
        End Set
    End Property

    'LOG Text ==============================
    Public ReadOnly Property LOG() As String
        Get
            LOG = mLOG
        End Get
    End Property

    'Error message ===========================
    Public ReadOnly Property ERR_MSG() As String
        Get
            ERR_MSG = mErrorComment
        End Get
    End Property

    Public Function TestPreparation() As Boolean

        'ﾒｲﾝCPU（ｽﾛｯﾄ1：SP7x）→KPをﾓﾆﾀする場合の前準備
        'KP内部ﾚｼﾞｽﾀがｱｸｾｽできるようIOﾊﾞｽﾒﾓﾘ空間にﾏｯﾋﾟﾝｸﾞしておく（？）
        'CPU検査時はｽﾛｯﾄ5が基本！

        TestPreparation = False

        Dim cMcm1 As cls120_SP_M_Command

        Dim slot As Integer
        If mSlotAddress = 0 Then
            slot = 5
        Else
            slot = mSlotAddress
        End If

        Dim WriteAddress As String = ""
        Dim WriteDATA As String = ""
        Dim mMessage As String = ""

        '検査ｽﾛｯﾄによって書き込むｱﾄﾞﾚｽが異なる為下記処理をする
        Dim Adres1 As String = (&H78 + (slot - 1) * &H80).ToString("X8")
        Dim Adres2 As String = (&H1A + (slot - 1) * &H80).ToString("X8")
        Dim Adres3 As String = (&H10 + (slot - 1) * &H80).ToString("X8")
        Dim Adres4 As String = (&H12 + (slot - 1) * &H80).ToString("X8")
        Dim Adres5 As String = (&H29 + (slot - 1) * &H80).ToString("X8")
        
        '==========================================================================================
        Call WriteLog("IOﾊﾞｽﾃｽﾄ前準備")

        cMcm1 = New cls120_SP_M_Command
        cMcm1.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm1.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm1.SlotNo = 1                         'CPU Slot番号
        cMcm1.UnitNo = 0                         'ユニット番号

        Try
            For i = 0 To 4
                Select Case i
                    Case 0
                        WriteAddress = Adres1 : WriteDATA = "0000" : mMessage = "ﾊﾟﾗﾒｰﾀ終了"
                    Case 1
                        WriteAddress = Adres2 : WriteDATA = "0080" : mMessage = "ｲﾆｼｬﾗｲｽﾞ終了"
                    Case 2
                        WriteAddress = Adres3 : WriteDATA = "0006" : mMessage = "上位ｳｩﾝﾄﾞｳｱﾄﾞﾚｽ設定"
                    Case 3
                        WriteAddress = Adres4 : WriteDATA = "2C18" : mMessage = "下位ｳｩﾝﾄﾞｳｱﾄﾞﾚｽ設定"
                    Case 4
                        WriteAddress = Adres5 : WriteDATA = "0007" : mMessage = "IOMEMﾎﾞｰﾚｰﾄｺﾝﾄﾛｰﾙﾚｼﾞｽﾀ設定"
                End Select

                If i <> 4 Then
                    Call cMcm1.IO_WRITE(1, 2, WriteAddress, WriteDATA)
                Else
                    Call cMcm1.IO_WRITE(0, 2, WriteAddress, WriteDATA)
                End If
                Call WriteLog(mMessage)
                Call WriteLog("ｱﾄﾞﾚｽ:" & WriteAddress & ",書き込みﾃﾞｰﾀ:" & WriteDATA)

            Next
            
            Call mSecWait(100)
            Call cMcm1.IO_READ(0, 1, Adres5)                 'KP_IOMEMﾎﾞｰﾚｰﾄｺﾝﾄﾛｰﾙﾚｼﾞｽﾀ（0x29）確認
            Call WriteLog("Send : " & "KP_IOMEMﾎﾞｰﾚｰﾄｺﾝﾄﾛｰﾙﾚｼﾞｽﾀ（0x29）（ｽﾛｯﾄ" & slot.ToString & "：" & Adres5 & "）読み込み")
            Call WriteLog("Recv : " & cMcm1.RecvText)
            If Mid(Strings.Right(cMcm1.RecvText, 4), 1, 2) <> "07" Then
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            mErrorComment = "IOﾊﾞｽ検査前準備ｴﾗｰ"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cMcm1 = Nothing
        End Try

        TestPreparation = True

    End Function

    Public Function SlotAddressTest() As Boolean

        'ｽﾛｯﾄ5の状態から検査をする
        'ｽﾛｯﾄ2,3,9の状態にしてｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀﾓﾆﾀでﾓﾆﾀできること

        SlotAddressTest = False

        Dim mReadAddress As String = ""

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'SLOTを確認
        Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽ検査")

        cBUS = New cls122_FAM3_IOBUS_Ctrl
        Try
            For i As Integer = 2 To 9

                Call cBUS.BUSMONITOR_TB04_CPU_SLOTChange(i)
                Call WriteLog("SLOT" & CStr(i) & "ｱﾄﾞﾚｽ確認")
                Call cBUS.IOBusTstSetReSetCpuXDFAIL()
                Call WaitTimer(5)

                '前準備
                mSlotAddress = i
                Call TestPreparation()

                Call mSecWait(100)

                Select Case i
                    Case 2
                        mRegSlotAddress = "00000080" : mRegIRQMon = "00000087" : mRegBRMon = "000000B0" : mRegMODFail = "000000A4" : mRegLFail = "00000097"
                    Case 3
                        mRegSlotAddress = "00000100" : mRegIRQMon = "00000107" : mRegBRMon = "00000130" : mRegMODFail = "00000124" : mRegLFail = "00000117"
                    Case 5
                        mRegSlotAddress = "00000200" : mRegIRQMon = "00000207" : mRegBRMon = "00000230" : mRegMODFail = "00000224" : mRegLFail = "00000217"
                    Case 9
                        mRegSlotAddress = "00000400" : mRegIRQMon = "00000407" : mRegBRMon = "00000430" : mRegMODFail = "00000424" : mRegLFail = "00000417"
                End Select

                Call cMcm.IO_READ(1, 1, mRegSlotAddress)
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegSlotAddress & "ﾘｰﾄﾞ（ｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀ0x00）")
                Call WriteLog("Recv : " & cMcm.RecvText)
                If Mid(Strings.Right(cMcm.RecvText, 4), 1, 1) <> CStr(i - 1) Then
                    mErrorComment = "SLOTｱﾄﾞﾚｽ認識異常：SLOT" & i.ToString
                    Throw New ApplicationException(mErrorComment)
                End If

                For j = 0 To 3
                    Select Case j
                        Case 0 : mReadAddress = mRegIRQMon
                        Case 1 : mReadAddress = mRegBRMon
                        Case 2 : mReadAddress = mRegMODFail
                        Case 3 : mReadAddress = mRegLFail
                    End Select

                    If j = 3 Then
                        Call cMcm.IO_READ(0, 1, mReadAddress)
                    Else
                        Call cMcm.IO_READ(1, 1, mReadAddress)
                    End If

                    Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                    Call WriteLog("Recv : " & cMcm.RecvText)

                    If j = 1 Then
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "80" Then
                            mErrorComment = "SLOTｱﾄﾞﾚｽにその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Else
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                            mErrorComment = "SLOTｱﾄﾞﾚｽにその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    End If

                Next j

                Call WriteLog("SLOT" & i.ToString & " OK")

                If i = 3 Then i = i + 5 '次はｽﾛｯﾄ9を検査
                
            Next i

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            'Call cBUS.BUSMONITOR_TB04_CPU_SLOTChange(5)
            'Call cBUS.IOBusTstSetReSetCpuXDFAIL()
            'Call WaitTimer(5)
            ''前準備
            'mSlotAddress = 5
            'Call TestPreparation()
            cBUS = Nothing
            cMcm = Nothing
        End Try

        cBUS = Nothing
        cMcm = Nothing

        SlotAddressTest = True

    End Function

    Public Function SlotAddressTestForOldSP() As Boolean

        'ｽﾛｯﾄ5の状態から検査をする
        'ｽﾛｯﾄ2,3,9の状態にしてｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀﾓﾆﾀでﾓﾆﾀできること

        SlotAddressTestForOldSP = False

        Dim mReadAddress As String = ""

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'SLOTを確認
        Call WriteLog("ｽﾛｯﾄｱﾄﾞﾚｽ検査")

        cBUS = New cls122_FAM3_IOBUS_Ctrl
        Try
            For i As Integer = 2 To 9

                Call cBUS.BUSMONITOR_TB04_CPU_SLOTChange(i)
                Call WriteLog("SLOT" & CStr(i) & "ｱﾄﾞﾚｽ確認")
                Call cBUS.IOBusTstSetReSetCpuXDFAIL()
                Call WaitTimer(5)

                '前準備
                mSlotAddress = i
                Call TestPreparation()

                Call mSecWait(100)

                Select Case i
                    Case 2
                        mRegSlotAddress = "00000080" : mRegIRQMon = "00000087" : mRegBRMon = "000000B0" : mRegMODFail = "000000A4" : mRegLFail = "00000097"
                    Case 3
                        mRegSlotAddress = "00000100" : mRegIRQMon = "00000107" : mRegBRMon = "00000130" : mRegMODFail = "00000124" : mRegLFail = "00000117"
                    Case 5
                        mRegSlotAddress = "00000200" : mRegIRQMon = "00000207" : mRegBRMon = "00000230" : mRegMODFail = "00000224" : mRegLFail = "00000217"
                    Case 9
                        mRegSlotAddress = "00000400" : mRegIRQMon = "00000407" : mRegBRMon = "00000430" : mRegMODFail = "00000424" : mRegLFail = "00000417"
                End Select

                Call cMcm.IO_READ(1, 1, mRegSlotAddress)
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegSlotAddress & "ﾘｰﾄﾞ（ｽﾛｯﾄｱﾄﾞﾚｽﾚｼﾞｽﾀ0x00）")
                Call WriteLog("Recv : " & cMcm.RecvText)
                If Mid(Strings.Right(cMcm.RecvText, 4), 1, 1) <> CStr(i - 1) Then
                    mErrorComment = "SLOTｱﾄﾞﾚｽ認識異常：SLOT" & i.ToString
                    Throw New ApplicationException(mErrorComment)
                End If

                For j = 0 To 3
                    Select Case j
                        Case 0 : mReadAddress = mRegIRQMon
                        Case 1 : mReadAddress = mRegBRMon
                        Case 2 : mReadAddress = mRegMODFail
                        Case 3 : mReadAddress = mRegLFail
                    End Select

                    If j = 3 Then
                        Call cMcm.IO_READ(0, 1, mReadAddress)
                    Else
                        Call cMcm.IO_READ(1, 1, mReadAddress)
                    End If

                    Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                    Call WriteLog("Recv : " & cMcm.RecvText)

                    If j = 1 Then
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "80" Then
                            mErrorComment = "SLOTｱﾄﾞﾚｽにその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    ElseIf j = 2 Then
                        If i = 2 Or i = 3 Then
                            If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "10" Or Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then

                            Else
                                mErrorComment = "SLOTｱﾄﾞﾚｽにその他ｼｮｰﾄの疑いあり！"
                                Throw New ApplicationException(mErrorComment)
                            End If
                        Else
                            If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                                mErrorComment = "SLOTｱﾄﾞﾚｽにその他ｼｮｰﾄの疑いあり！"
                                Throw New ApplicationException(mErrorComment)
                            End If
                        End If
                    Else
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                            mErrorComment = "SLOTｱﾄﾞﾚｽにその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    End If

                Next j

                Call WriteLog("SLOT" & i.ToString & " OK")

                If i = 3 Then i = i + 5 '次はｽﾛｯﾄ9を検査

            Next i

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            cBUS = Nothing
            cMcm = Nothing
        End Try

        cBUS = Nothing
        cMcm = Nothing

        SlotAddressTestForOldSP = True

    End Function

    Public Function XACSTest() As Boolean
        'XACS検査(CPU)

        XACSTest = False

        Dim mintUnit As Integer = 0      'DUT検査ユニット番号
        Dim mintSlot As Integer = 6      'DUT検査スロット番号
        Dim mintXD6407 As Integer = 7    'BM128 ＤＩスロット番号
        Dim mintYD6408 As Integer = 8    'BM128 ＤＯスロット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl
        'IOバスアクセス
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd
        cMcm.RemPort = 12290                           'LE バイナリ形式時は 12290 にする。

        cMcm.SlotNo = 1        'CPU Slot番号
        cMcm.UnitNo = 0        'ユニット番号

        Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        'Y33->0　スルーする。XACSされない状態で
        'X33->0　XASCネゲートならＯＫ。
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 0)
        Call mSecWait(100)
        'ラッチ解除
        'Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)
        'Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 1)
        'Call IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)
        Dim startime As Double = System.Environment.TickCount
        Do
            Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)
            If (System.Environment.TickCount - startime) / 1000 > 1 Then Exit Do '1秒間XACSが0であること      v0.01
        Loop Until cBUS.RCV_DATA <> "0"

        If cBUS.RCV_DATA <> "0" Then
            cBUS = Nothing : cMcm = Nothing : Exit Function
        End If

        Call WriteLog("1秒間XACSﾈｹﾞｰﾄ確認OK")

        'Y33->1　ラッチする。
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 1)
        Call WriteLog("XACSﾓﾆﾀﾗｯﾁﾓｰﾄﾞｾｯﾄ")

        Call cMcm.IO_READ(0, 4, "00000228")      'ｽﾛｯﾄ5にMｺﾏﾝﾄﾞでIOﾊﾞｽｱｸｾｽ
        Call WriteLog("ｽﾛｯﾄ5にIOﾊﾞｽｱｸｾｽ実行")

        cMcm = Nothing

        'X33->1　XASCアサートならＯＫ。
        Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        If cBUS.RCV_DATA <> "1" Then
            cBUS = Nothing : cMcm = Nothing : Exit Function
        End If

        Call WriteLog("XACSｱｻｰﾄ確認OK")

        'Y33->0　スルーする(信号線が繋がっている場合アクセス後にクリア)
        '被検査モジュールに何らかのアクセス
        'X33->0　ならＯＫ。
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 33, 0)
        Call cBUS.IO_READ_INRELAY_POINT(mintUnit, mintXD6407, 33)

        'ラッチ解除
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 1)
        Call cBUS.IO_WRITE_OUTRELAY_POINT(mintUnit, mintYD6408, 34, 0)

        XACSTest = True
        cBUS = Nothing : cMcm = Nothing

    End Function

    Public Function XBRTest() As Boolean
        'ﾒｲﾝCPU（SP7x）にてXBR0～3を順次上げ下げして
        'IOﾊﾞｽﾚｼﾞｽﾀにてXBRの変化をﾓﾆﾀできること

        XBRTest = False

        Dim mReadAddress As String = ""

        'XBRﾃｽﾄ
        Call WriteLog("XBR検査")
        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                     'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Try
            For i As Integer = 3 To 0 Step -1
                Call WriteLog("XBR" & CStr(i) & "ｾｯﾄ")

                Select Case i
                    Case 3 : Call cMcm.MEM_WRITE(1, "3880070B", "08")
                    Case 2 : Call cMcm.MEM_WRITE(1, "3880070B", "04")
                    Case 1 : Call cMcm.MEM_WRITE(1, "3880070B", "02")
                    Case 0 : Call cMcm.MEM_WRITE(1, "3880070B", "01")
                End Select
                Call mSecWait(100)

                Call WriteLog("XBR" & CStr(i) & "確認")
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegBRMon & "ﾘｰﾄﾞ（0x30のｽﾛｯﾄ5BRﾚｼﾞｽﾀ）")
                Call cMcm.IO_READ(1, 1, mRegBRMon)
                Call WriteLog("Recv : " & cMcm.RecvText)
                Select Case Mid(Strings.Right(cMcm.RecvText, 4), 1, 1)
                    Case "8"
                        If i <> 3 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case "4"
                        If i <> 2 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case "2"
                        If i <> 1 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case "1"
                        If i <> 0 Then
                            mErrorComment = "XBR認識異常：XBR" & i.ToString
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Case Else
                        mErrorComment = "XBR認識異常：XBR"
                        Throw New ApplicationException(mErrorComment)
                End Select

                For j = 0 To 3
                    Select Case j
                        Case 0 : mReadAddress = mRegIRQMon
                        Case 1 : mReadAddress = mRegSlotAddress
                        Case 2 : mReadAddress = mRegMODFail
                        Case 3 : mReadAddress = mRegLFail
                    End Select

                    If j = 3 Then
                        Call cMcm.IO_READ(0, 1, mReadAddress)
                    Else
                        Call cMcm.IO_READ(1, 1, mReadAddress)
                    End If

                    Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                    Call WriteLog("Recv : " & cMcm.RecvText)

                    If j = 1 Then
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
                            mErrorComment = "XBRとその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    Else
                        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                            mErrorComment = "XBRとその他ｼｮｰﾄの疑いあり！"
                            Throw New ApplicationException(mErrorComment)
                        End If
                    End If

                Next j

                Call WriteLog("XBR" & i.ToString & " OK")

            Next i

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cMcm.MEM_WRITE(1, "3880070B", "08")
            'Call cBUS.BUSMONITOR_TB04_CPU_XBRChange(5)
            cBUS = Nothing
            cMcm = Nothing
        End Try

        cBUS = Nothing
        cMcm = Nothing

        XBRTest = True

    End Function

    Public Function XIRQTest() As Boolean

        'ﾒｲﾝCPU（SP7x）のXIRQをﾏｽｸしてﾊﾞｽﾓﾆﾀ側からXIRQ0～3を上げ下げして
        'IOﾊﾞｽﾚｼﾞｽﾀにてXIRQの変化がﾓﾆﾀできること
        XIRQTest = False

        Dim mResStr As String = ""
        Dim mReadAddress As String = ""

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        'ﾒｲﾝCPUのXIRQﾏｽｸ
        Call WriteLog("ﾒｲﾝCPUのXIRQﾏｽｸ")
        Call cMcm.MEM_WRITE(1, "3880021E", "00")

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Try
            For i As Integer = 3 To 0 Step -1
                Call cBUS.BUSMONITOR_TB04_CPU_XIRQChange(i)
                Call WriteLog("XIRQ" & CStr(i) & "確認")
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegIRQMon & "ﾘｰﾄﾞ（0x07のｽﾛｯﾄ5IRQﾚｼﾞｽﾀ）")
                Call cMcm.IO_READ(1, 1, mRegIRQMon)
                Call WriteLog("Recv : " & cMcm.RecvText)

                Select Case i
                    Case 0 : mResStr = "8"
                    Case 1 : mResStr = "4"
                    Case 2 : mResStr = "2"
                    Case 3 : mResStr = "1"
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

                    Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                    Call WriteLog("Recv : " & cMcm.RecvText)

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

                Call WriteLog("XIRQ" & i.ToString & " OK")

            Next
        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XIRQ異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.BUSMONITOR_TB04_CPU_XIRQChange(5)
            'ﾒｲﾝCPUのXIRQﾏｽｸ解除
            Call WriteLog("ﾒｲﾝCPUのXIRQﾏｽｸ解除")
            Call cMcm.MEM_WRITE(1, "3880021E", "88")
            cBUS = Nothing
            cMcm = Nothing
        End Try

        Call WriteLog("XIRQ：OK")

        XIRQTest = True

    End Function

    Public Function XLFAILTest() As Boolean

        XLFAILTest = False

        Dim mReadAddress As String = ""
        Dim rData As Integer

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call cMcm.IO_READ(0, 1, mRegLFail)
        Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegLFail & "ﾘｰﾄﾞ（0x17のｽﾛｯﾄ5XLFAILﾚｼﾞｽﾀ）")
        Call WriteLog("Recv : " & cMcm.RecvText)
        rData = CByte("&H" & Mid(Strings.Right(cMcm.RecvText, 4), 1, 2))
        If (&H2 And rData) <> 0 Then
            mErrorComment = "XLFAIL異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        'XLFAILｾｯﾄ
        If cBUS.IOBusTstSetCpuXLFAIL() = False Then
            Call cBUS.IOBusTstReSetCpuXLFAIL()
            mErrorComment = "XLFAILｾｯﾄ異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog("XLFAILｾｯﾄ")

        Try
            Call cMcm.IO_READ(0, 1, mRegLFail)
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegLFail & "ﾘｰﾄﾞ（0x17のｽﾛｯﾄ5XLFAILﾚｼﾞｽﾀ）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            rData = CByte("&H" & Mid(Strings.Right(cMcm.RecvText, 4), 1, 2))
            If (&H2 And rData) <> 2 Then
                mErrorComment = "XLFAIL異常"
                Throw New ApplicationException(mErrorComment)
            End If

            For j = 0 To 3
                Select Case j
                    Case 0 : mReadAddress = mRegBRMon
                    Case 1 : mReadAddress = mRegSlotAddress
                    Case 2 : mReadAddress = mRegMODFail
                    Case 3 : mReadAddress = mRegIRQMon
                End Select

                Call cMcm.IO_READ(1, 1, mReadAddress)

                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                Call WriteLog("Recv : " & cMcm.RecvText)

                If j = 1 Then
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
                        mErrorComment = "XLFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                ElseIf j = 0 Then
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "80" Then
                        mErrorComment = "XLFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                Else
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                        mErrorComment = "XLFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                End If

            Next j
        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXLFAIL()
            Call WriteLog("XLFAILﾘｾｯﾄOK")
            cMcm = Nothing
            cBUS = Nothing
        End Try

        Call WriteLog("XLFAIL：OK")
        XLFAILTest = True

    End Function

    Public Function XPFAILTest() As Boolean

        Dim mReadAddress As String = ""

        XPFAILTest = False

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call WriteLog("ﾒｲﾝCPUのXPFAILﾏｽｸ")
        '事前準備1
        Call cMcm.MEM_WRITE(1, "38800013", "00")
        '事前準備2
        Call cMcm.MEM_WRITE(1, "38800011", "00")

        Try
            Call cMcm.IO_READ(1, 1, mRegMODFail)
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegMODFail & "ﾘｰﾄﾞ（0x24のｽﾛｯﾄ5）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                mErrorComment = "XPFAIL異常"
                Throw New ApplicationException(mErrorComment)
            End If

            'XPFAILｾｯﾄ
            If cBUS.IOBusTstSetCpuXPFAIL() = False Then
                Call cBUS.IOBusTstReSetCpuXPFAIL()
                mErrorComment = "XPFAILｾｯﾄ異常"
                Throw New ApplicationException(mErrorComment)
            End If

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Call cMcm.MEM_WRITE(1, "38800011", "80")
            cMcm = Nothing
            cBUS = Nothing
            Throw New ApplicationException(mErrorComment)
        End Try

        Call WriteLog("XPFAILｾｯﾄ")

        Try
            Call cMcm.IO_READ(1, 1, mRegMODFail)
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegMODFail & "ﾘｰﾄﾞ（0x24のｽﾛｯﾄ5）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
                mErrorComment = "XPFAIL認識異常"
                Throw New ApplicationException(mErrorComment)
            End If

            For j = 0 To 3
                Select Case j
                    Case 0 : mReadAddress = mRegBRMon
                    Case 1 : mReadAddress = mRegSlotAddress
                    Case 2 : mReadAddress = mRegLFail
                    Case 3 : mReadAddress = mRegIRQMon
                End Select

                If j = 2 Then
                    Call cMcm.IO_READ(0, 1, mReadAddress)
                Else
                    Call cMcm.IO_READ(1, 1, mReadAddress)
                End If
                
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                Call WriteLog("Recv : " & cMcm.RecvText)

                If j = 1 Then
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
                        mErrorComment = "XPFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                ElseIf j = 0 Then
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "80" Then
                        mErrorComment = "XPFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                Else
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                        mErrorComment = "XPFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                End If

            Next j

        Catch ex As Exception
            If mErrorComment = "" Then mErrorComment = "XPFAIL認識異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXPFAIL()
            Call WriteLog("XPFAILﾘｾｯﾄOK")
            Call WriteLog("ﾒｲﾝCPUのXPFAILﾏｽｸ解除")
            Call cMcm.MEM_WRITE(1, "38800011", "80")
            cMcm = Nothing
            cBUS = Nothing
        End Try

        Call WriteLog("XPFAIL：OK")

        XPFAILTest = True

    End Function

    Public Function XSFAILTest() As Boolean

        XSFAILTest = False

        Dim mReadAddress As String = ""

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                     'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call cBUS.IOBusTstReSetCpuXSFAIL()

        Call cMcm.IO_READ(1, 1, mRegMODFail)
        Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegMODFail & "ﾘｰﾄﾞ（0x24のｽﾛｯﾄ5）")
        Call WriteLog("Recv : " & cMcm.RecvText)
        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
            mErrorComment = "XSFAIL異常"
            Throw New ApplicationException(mErrorComment)
        End If

        'XSFAILｾｯﾄ
        If cBUS.IOBusTstSetCpuXSFAIL() = False Then
            Call cBUS.IOBusTstReSetCpuXSFAIL()
            mErrorComment = "XSFAILｾｯﾄ異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog("XSFAILｾｯﾄ")

        Try
            Call cMcm.IO_READ(1, 1, mRegMODFail)
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegMODFail & "ﾘｰﾄﾞ（0x24のｽﾛｯﾄ5）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "10" Then
                mErrorComment = "XSFAIL認識異常"
                Throw New ApplicationException(mErrorComment)
            End If

            For j = 0 To 3
                Select Case j
                    Case 0 : mReadAddress = mRegBRMon
                    Case 1 : mReadAddress = mRegSlotAddress
                    Case 2 : mReadAddress = mRegLFail
                    Case 3 : mReadAddress = mRegIRQMon
                End Select

                If j = 2 Then
                    Call cMcm.IO_READ(0, 1, mReadAddress)
                Else
                    Call cMcm.IO_READ(1, 1, mReadAddress)
                End If

                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                Call WriteLog("Recv : " & cMcm.RecvText)

                If j = 1 Then
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
                        mErrorComment = "XSFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                ElseIf j = 0 Then
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "80" Then
                        mErrorComment = "XSFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                Else
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                        mErrorComment = "XSFAILとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                End If

            Next j

        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXSFAIL()
            Call WriteLog("XSFAILﾘｾｯﾄOK")
            cMcm = Nothing
            cBUS = Nothing
        End Try

        XSFAILTest = True
        Call WriteLog("XSFAIL：OK")

    End Function

    Public Function XBCLRTest() As Boolean

        XBCLRTest = False

        Dim mReadAddress As String = ""

        cMcm = New cls120_SP_M_Command
        cMcm.IPAdd = mIPAdd                      'CPU用LE経由
        cMcm.RemPort = 12290                    'LE バイナリ形式時は 12290 にする。
        cMcm.SlotNo = 1                         'CPU Slot番号
        cMcm.UnitNo = 0                         'ユニット番号

        cBUS = New cls122_FAM3_IOBUS_Ctrl

        Call cMcm.IO_READ(1, 1, mRegBRMon)
        Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegBRMon & "ﾘｰﾄﾞ（0x30のｽﾛｯﾄ5XBRﾚｼﾞｽﾀ）")
        Call WriteLog("Recv : " & cMcm.RecvText)
        If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "80" Then
            mErrorComment = "XBCLR異常"
            Throw New ApplicationException(mErrorComment)
        End If

        'XBCLRｾｯﾄ
        If cBUS.IOBusTstSetCpuXBCLR() = False Then
            Call cBUS.IOBusTstReSetCpuXBCLR()
            mErrorComment = "XBCLRｾｯﾄ異常"
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        End If

        Call WriteLog("XBCLRｾｯﾄ")

        Try
            Call cMcm.IO_READ(1, 1, mRegBRMon)
            Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mRegBRMon & "ﾘｰﾄﾞ（0x30のｽﾛｯﾄ5XBRﾚｼﾞｽﾀ）")
            Call WriteLog("Recv : " & cMcm.RecvText)
            If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "88" Then
                mErrorComment = "XBCLR認識異常"
                Throw New ApplicationException(mErrorComment)
            End If

            For j = 0 To 3
                Select Case j
                    Case 0 : mReadAddress = mRegLFail
                    Case 1 : mReadAddress = mRegSlotAddress
                    Case 2 : mReadAddress = mRegMODFail
                    Case 3 : mReadAddress = mRegIRQMon
                End Select

                If j = 0 Then
                    Call cMcm.IO_READ(0, 1, mReadAddress)
                Else
                    Call cMcm.IO_READ(1, 1, mReadAddress)
                End If
                
                Call WriteLog("Send : " & "ｱﾄﾞﾚｽ" & mReadAddress & "ﾘｰﾄﾞ")
                Call WriteLog("Recv : " & cMcm.RecvText)

                If j = 1 Then
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "40" Then
                        mErrorComment = "XBCLRとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                Else
                    If Mid(Strings.Right(cMcm.RecvText, 4), 1, 2) <> "00" Then
                        mErrorComment = "XBCLRとその他ｼｮｰﾄの疑いあり！"
                        Throw New ApplicationException(mErrorComment)
                    End If
                End If

            Next j
        Catch ex As Exception
            Call WriteLog(mErrorComment)
            Throw New ApplicationException(mErrorComment)
        Finally
            Call cBUS.IOBusTstReSetCpuXBCLR()
            Call WriteLog("XBCLRﾘｾｯﾄOK")
            cMcm = Nothing
            cBUS = Nothing
        End Try

        Call WriteLog("XBCLR：OK")
        XBCLRTest = True

    End Function


    Private Sub WriteLog(ByVal mComment As String)
        mLOG &= mComment & vbCrLf
    End Sub

    Private Sub mSecWait(ByVal mSec As Integer)
        System.Threading.Thread.Sleep(mSec)         'mSec秒のWait
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

End Class
