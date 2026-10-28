Public Class cls034_LabelPrint

    Dim cLbl As cls011_Excel_Operation
    Dim cADO As cls010_MDB_Operation

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"
    Dim mErrorComment As String

    Dim mPrintForm As String

    'Property =====================================================================================
    'Print Form Model
    Public WriteOnly Property PrintForm() As String
        Set(ByVal rp As String)
            mPrintForm = rp
        End Set
    End Property

    'Error message --------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public Sub LabelPrint(ByVal dbName As String, ByVal rDTkey As String)

        Dim wTime As Double = 0

        Dim tbName As String
        Dim sSQL As String
        Dim fName() As String = {""}
        Dim sData() As String = {""}
        Dim rField As String
        'Dim kField As String
        'Dim rCount As Integer

        Dim sModel As String
        Dim sXJNO As String
        Dim sHRev As String
        Dim sSRev As String
        Dim sMac1 As String
        Dim sMac2 As String
        Dim sEquipmentNo As String
        Dim sUSB As String
        Dim sIEEE As String
        Dim sLSer As String
        Dim sDate As String
        Dim sInsModel As String

        Dim mPrintUL As String = ""
        Dim mPrintCE As String = ""
        Dim mPrintWEEE As String = ""
        Dim mPrintCTick As String = ""
        Dim mPrintRoHS As String = ""
        Dim mPrintModuleName As String = ""
        Dim mPrintModel As String
        Dim mPrintSUFFIX1 As String
        Dim mPrintSUFFIX2 As String = ""

        'Dim mPrintKC As String
        Dim mPrintKCNo As String
        'Dim mPrintKCModel As String

        Dim mReadSUFFIX2 As String

        Dim mRModel As String
        Dim mSupply(1) As String
        Dim mInput(2) As String
        Dim mOutput(3) As String

        Dim mTitle(5) As String
        Dim mString(5) As String

        Dim rDate As Date
        Dim nwYear As String
        Dim nwMonth As String
        Dim nwDay As String

        Dim mPrinterName As String

        '==========================================================================================

        '縦ﾀｲﾄﾙの読み込み
        Try
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = mAppPath & "410 Print Form\" & "主銘板印字タイトル.xls"
            cADO.TB_NAME = "[Sheet1$]"
            rField = "title1,title2,title3,title4,title5,title6"
            sSQL = "select " & rField & " from " & "[Sheet1$]" & " where FORM like '" & mPrintForm & "'"
            cADO.DB_SQL = sSQL
            Call cADO.DB_READ_EXCEL_JET(fName, sData)

            If cADO.DB_DATA_COUNT = 0 Then
                mErrorComment = "主銘板に印字するタイトルを取得できませんでした。"
                Throw New ApplicationException(mErrorComment)
            End If

            mTitle = sData

        Catch ex As Exception
            If mErrorComment = "" Then
                mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
                mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
                'Call WriteLog(False, mErrorComment)
            End If
            Throw New ApplicationException(mErrorComment)
        Finally
            cADO = Nothing
        End Try

        Try
            '印字データの読込み -----------------------------
            cADO = New cls010_MDB_Operation
            cADO.DB_PATH = dbName
            tbName = "tblResult"
            cADO.TB_NAME = tbName
            rField = "Mscode,XJNo,ProductHardRev,ProductSoftRev,MacAddress,SerialNo,EquipmentNo,StartDateTime,InspectionModel,MacAddress2,USBAddress,IEEE1394Address"
            'rField = "Mscode,XJNo,ProductHardRev,ProductSoftRev,MacAddress,SerialNo,EquipmentNo,StartDateTime,InspectionModel"     'FD時
            'rField = "Mscode,XJNo,ProductHardRev,ProductSoftRev,MacAddress,LineSerial,EquipmentNo,StartDateTime"    'debug時

            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE " & "StartDateTime LIKE '" & rDTkey & "'"
            sSQL &= " ORDER BY StartDateTime DESC;"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)

            If cADO.DB_DATA_COUNT = 0 Then
                mErrorComment = "主銘板に印字するデータを取得できませんでした。"
                Throw New ApplicationException(mErrorComment)
            ElseIf sData(0) = "" Or sData(2) = "" Or sData(3) = "" Or sData(5) = "" Or sData(6) = "" Then
                mErrorComment = "主銘板に印字するデータに不備があります。"
                Throw New ApplicationException(mErrorComment)
            End If

            sModel = sData(0) : sXJNO = sData(1) : sHRev = sData(2) : sSRev = sData(3) : sMac1 = sData(4)
            sLSer = sData(5) : sEquipmentNo = sData(6) : sInsModel = sData(8)
            sMac2 = sData(9) : sUSB = sData(10) : sIEEE = sData(11)

            rDate = DateValue(sData(7))
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
            'rField = "ModuleName,Model6,Suffix1,RoHSMark,ULMark,CEMark,WEEEMark,CTickMark,Model,Supply,Input1,Input2,Input3,Output1,Output2,Output3,Output4,Suffix2"        '
            rField = "ModuleName,Model6,Suffix1,RoHSMark,ULMark,CEMark,WEEEMark,CTickMark,Model,Supply,Input1,Input2,Input3,Output1,Output2,Output3,Output4,Suffix2,"
            rField &= "KCNO,Supply2"
            sSQL = "SELECT " & rField & " FROM " & tbName & " WHERE Model = '" & sInsModel & "'"
            cADO.DB_SQL = sSQL

            Call cADO.DB_READ_JET(fName, sData)
            cADO = Nothing

            mPrintModuleName = sData(0) : mPrintModel = sData(1) : mPrintSUFFIX1 = sData(2)
            mPrintRoHS = sData(3) : mPrintUL = sData(4) : mPrintCE = sData(5) : mPrintWEEE = sData(6)
            mPrintCTick = sData(7) : mRModel = sData(8)
            mSupply(0) = sData(9) : mInput(0) = sData(10) : mInput(1) = sData(11) : mInput(2) = sData(12)
            mOutput(0) = sData(13) : mOutput(1) = sData(14) : mOutput(2) = sData(15) : mOutput(3) = sData(16)
            mReadSUFFIX2 = sData(17)

            mPrintKCNo = sData(18)
            mSupply(1) = sData(19)

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
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            cADO = Nothing

        End Try

        'suffix2の決定
        Dim sfx2() As String = Split(mReadSUFFIX2, "，")       '全角カンマ区切り
        Dim OrderModel() As String = Split(sModel, "/")
        Dim XJModel() As String = Split(sXJNO, "/")

        For i = 0 To UBound(sfx2)
            For j = 0 To UBound(OrderModel)
                If sfx2(i) = "/" & OrderModel(j) Then mPrintSUFFIX2 &= sfx2(i) : GoTo looploop
            Next
            For k = 0 To UBound(XJModel)
                If sfx2(i) = "/" & XJModel(k) Then mPrintSUFFIX2 &= sfx2(i)
            Next
looploop:
        Next

        '印字データの決定
        For i = 0 To UBound(mTitle)
            Try
                Select Case mTitle(i)
                    Case "SUPPLY" : mString(i) = mSupply(0)
                    Case "SUPPLY2" : mString(i) = mSupply(1) : mTitle(i) = ""
                    Case "INPUT1" : mString(i) = mInput(0) : mTitle(i) = "INPUT"
                    Case "INPUT2" : mString(i) = mInput(1) : mTitle(i) = ""
                    Case "INPUT3" : mString(i) = mInput(2) : mTitle(i) = ""
                    Case "OUTPUT1" : mString(i) = mOutput(0) : mTitle(i) = "OUTPUT"
                    Case "OUTPUT2" : mString(i) = mOutput(1) : mTitle(i) = ""
                    Case "OUTPUT3" : mString(i) = mOutput(2) : mTitle(i) = ""
                    Case "OUTPUT4" : mString(i) = mOutput(3) : mTitle(i) = ""
                    Case "MAC ID"
                        If sMac1 = "" Then
                            mErrorComment = "主銘板に印字するデータに不備があります。MAC IDなし"
                            Throw New ApplicationException(mErrorComment)
                        Else
                            mString(i) = "'" & sMac1
                        End If
                    Case "MAC ID1"
                        If sMac1 = "" Then
                            mErrorComment = "主銘板に印字するデータに不備があります。MAC IDなし"
                            Throw New ApplicationException(mErrorComment)
                        Else
                            mString(i) = "'" & sMac1
                        End If
                    Case "MAC ID2"
                        If sMac2 = "" Then
                            mErrorComment = "主銘板に印字するデータに不備があります。MAC ID2なし"
                            Throw New ApplicationException(mErrorComment)
                        Else
                            mString(i) = "'" & sMac2
                        End If
                    Case "USB"
                        If sUSB = "" Then
                            mErrorComment = "主銘板に印字するデータに不備があります。USBなし"
                            Throw New ApplicationException(mErrorComment)
                        Else
                            mString(i) = "'" & sUSB
                        End If
                    Case "IEEE1394"
                        If sIEEE = "" Then
                            mErrorComment = "主銘板に印字するデータに不備があります。IEEE1394なし"
                            Throw New ApplicationException(mErrorComment)
                        Else
                            mString(i) = "'" & sIEEE
                        End If
                    Case "US PAT" : mString(i) = "'5805442,6282454"
                    Case "" : mString(i) = ""
                End Select
            Catch ex As Exception
                Throw New ApplicationException(mErrorComment)
            End Try

        Next

        Try
            cLbl = New cls011_Excel_Operation
            With cLbl
                If mPrintKCNo <> "" Then
                    .FILE_PATH = mAppPath & "410 Print Form\" & "FA-M3 Master for KC.xls"
                Else
                    .FILE_PATH = mAppPath & "410 Print Form\" & "FA-M3 Master.xls"
                End If

                .SHEET_NAME = "Sheet1"
                .EXCEL_START()

                .CELL_DATA_SET("MODULENAME", mPrintModuleName)
                .CELL_DATA_SET("MODEL", mPrintModel)
                .CELL_DATA_SET("SUFFIX1", mPrintSUFFIX1)
                .CELL_DATA_SET("SUFFIX2", mPrintSUFFIX2)

                .CELL_DATA_SET("REV", sSRev & ":" & sHRev)

                For i = 0 To 5
                    .CELL_DATA_SET("title" & CStr(i + 1), mTitle(i))
                    .CELL_DATA_SET("str" & CStr(i + 1), mString(i))
                Next

                .CELL_DATA_SET("DATE", sDate)
                .CELL_DATA_SET("FAM3SERIAL", sLSer & " " & "0" & sEquipmentNo)

                .CELL_DATA_SET("MODELH", mPrintModel & mPrintSUFFIX1 & mPrintSUFFIX2)
                .CELL_DATA_SET("FAM3SERIALH", sLSer & " " & "0" & sEquipmentNo)

                Dim bc1 As DotNetBarcode = New DotNetBarcode
                bc1.Type = DotNetBarcode.Types.QRCode
                bc1.QRCopyToClipboard(sLSer & sEquipmentNo & nwYear & nwMonth & nwDay & Mid(mRModel, 1, 6) & Mid(mRModel, 8, mRModel.Length - 7), 7)
                .PICTURE_QRPASTE_CONTROL()
                '.CELL_DATA_SET("BARCODE", sLSer & sEquipmentNo & nwYear & nwMonth & nwDay & Mid(mRModel, 1, 6) & Mid(mRModel, 8, mRModel.Length - 7))

                '.BARCODE_SHOW("BARCODE", 120 + 180 - 120, 285, 135, 27)
                '.BARCODE_SHOW("BARCODE", 120 + 180 - 140, 285, 155, 27)
                If mPrintRoHS = "True" Then .CELL_DATA_SET("ROHS", "'RoHS")

                If InStr(sModel, "/C01") <> 0 Then
                    .PICTURE_VISIBLE_CONTROL("【YOKOGAWA】", False)
                    .PICTURE_VISIBLE_CONTROL("【ユーシン】", True)
                Else
                    .PICTURE_VISIBLE_CONTROL("【YOKOGAWA】", True)
                    .PICTURE_VISIBLE_CONTROL("【ユーシン】", False)
                End If

                If mPrintUL = "False" Then .PICTURE_VISIBLE_CONTROL("UL", False)
                If mPrintCE = "False" Then .PICTURE_VISIBLE_CONTROL("CE", False)
                If mPrintWEEE = "False" Then .PICTURE_VISIBLE_CONTROL("WEEE", False)
                If mPrintCTick = "False" Then .PICTURE_VISIBLE_CONTROL("C-Tick", False)

                If mPrintKCNo <> "" Then
                    .CELL_DATA_SET("KCNO", mPrintKCNo)
                    'Dim kcmodel() As String = Split(mRModel, "/")
                    'If UBound(kcmodel) = 0 Then
                    '    .CELL_DATA_SET("KCModel", mRModel)
                    'ElseIf UBound(kcmodel) > 0 Then
                    '    If kcmodel(1) <> "L1" Then
                    '        .CELL_DATA_SET("KCModel", kcmodel(0))
                    '    Else
                    '        .CELL_DATA_SET("KCModel", mRModel)
                    '    End If
                    'End If
                    '.CELL_DATA_SET("KCModel", mPrintKCModel)
                End If

                .PRINT_EXCEL(mPrinterName)

            End With

        Catch ex As Exception
            mErrorComment = Mid(ex.ToString, InStr(ex.ToString, ":") + 1, Len(ex.ToString))
            mErrorComment = Mid(mErrorComment, 1, InStr(mErrorComment, vbCr) - 1)
            'Call WriteLog(False, mErrorComment)
            Throw New ApplicationException(mErrorComment)

        Finally
            cLbl.CLOSE_EXCEL()
            cLbl = Nothing

        End Try

    End Sub

End Class
