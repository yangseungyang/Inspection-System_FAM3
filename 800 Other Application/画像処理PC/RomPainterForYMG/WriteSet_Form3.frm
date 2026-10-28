VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form WriteSet_Form3 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Write"
   ClientHeight    =   5970
   ClientLeft      =   6450
   ClientTop       =   4275
   ClientWidth     =   4365
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   4365
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Frame Frame1 
      Caption         =   "書込みｵﾌﾟｼｮﾝ"
      Height          =   3735
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   3855
      Begin VB.Frame Frame2 
         Caption         =   "ｴﾝﾃﾞｨｱﾝ"
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   2880
         Width           =   3255
         Begin VB.OptionButton EndianSelect 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   31
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton EndianSelect 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   30
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label15 
            Caption         =   "Little"
            Height          =   255
            Left            =   2280
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Big"
            Height          =   255
            Left            =   840
            TabIndex        =   32
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   3
         Left            =   1800
         OLEDropMode     =   1  '手動
         TabIndex        =   28
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   2
         Left            =   1800
         OLEDropMode     =   1  '手動
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   1
         Left            =   1800
         MaxLength       =   12
         TabIndex        =   19
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   420
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   2325
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   420
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   270
         Width           =   255
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   0
         Left            =   1800
         MaxLength       =   9
         TabIndex        =   14
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Height          =   420
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   1965
         Width           =   255
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   720
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label14 
         Caption         =   "基本ﾌﾞｰﾄ"
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   750
         Width           =   615
      End
      Begin VB.Line Line3 
         X1              =   1920
         X2              =   1920
         Y1              =   720
         Y2              =   840
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   1920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label FirmSum_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   2760
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.Label SystemSum_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   2760
         TabIndex        =   23
         Top             =   1560
         Width           =   735
      End
      Begin VB.Line Line5 
         X1              =   2040
         X2              =   1920
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line2 
         X1              =   1920
         X2              =   1920
         Y1              =   1560
         Y2              =   1680
      End
      Begin VB.Label Label13 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   1590
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "OSﾌﾞｰﾄ"
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "MACｱﾄﾞﾚｽ"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   2430
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "ﾓｼﾞｭｰﾙ型名"
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   2070
         Width           =   1215
      End
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   525
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   825
      Width           =   135
   End
End
Attribute VB_Name = "WriteSet_Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  WriteSet_Form3                                      *
'*                                                      *
'*      作成日  ：2007/06/18                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM書込み ｵﾌﾟｼｮﾝ画面（RTOS用）        *
'*                                                      *
'********************************************************

'Option Explicit

'---------------------------------------------
'外部API関数使用宣言
'---------------------------------------------
'---ｷｬﾚｯﾄを消す
Private Declare Function HideCaret Lib "user32.dll" (ByVal hwnd As Long) As Long


'--------------------------------------------------------
'
'   ROM書込みｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    Dim i As Integer
    
    '選択ﾌｧｲﾙのﾊﾟｽとｻｲｽﾞを保存しておく(ｷｬﾝｾﾙ時元に戻すため)
    PassForCancel = SelectedFilePass
    SizeForCancel = FileSize
    
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    SlotNo_lbl.Caption = SelectedSlot                                   'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名
        
    If Main_Form.FileName_txt.Text = "" Then
        For i = 0 To 3
            OptionText(i).BackColor = &H8000000B
        Next
    Else
        If FileSize < 65536 Then    '基本ﾌﾞｰﾄ
            Check1(2).Value = vbChecked
            OptionText(2).Text = Main_Form.FileName_txt.Text        '書込みﾌｧｲﾙ名
            FirmSum_lbl.Caption = Main_Form.SumValue_lbl.Caption    '計算ｻﾑ値
        Else                        'OSﾌﾞｰﾄ
            Check1(3).Value = vbChecked
            OptionText(3).Text = Main_Form.FileName_txt.Text        '書込みﾌｧｲﾙ名
            SystemSum_lbl.Caption = Main_Form.SumValue_lbl.Caption  '計算ｻﾑ値
            SystemFilePass = SelectedFilePass                       'OSﾌﾞｰﾄ用変数に代入
            SystemFileSize = FileSize                               'OSﾌﾞｰﾄ用変数に代入
        End If
        For i = 0 To 1
            OptionText(i).BackColor = &H8000000B
        Next
    End If
    
    EndianSelect(0).Value = True
    Endian = BIG

    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    WriteAddress = 0
    WriteSize = 0
    Endian = BIG
    SaveFilePass = SelectedFilePass         '元のﾌｧｲﾙのﾊﾟｽを保存しておく
    
    'ROM書込み用ﾌﾗｸﾞ初期化
    WriteMode = Normal_Write
    WriteFirm = DISABLE
    WriteSystem = DISABLE
    WriteName = DISABLE
    WriteMAC = DISABLE

End Sub


'--------------------------------------------------------
'
'   書込みｵﾌﾟｼｮﾝ
'
'--------------------------------------------------------
'--------------------------------------------------------
'   ﾁｪｯｸﾎﾞﾀﾝ
'--------------------------------------------------------
Private Sub Check1_Click(Index As Integer)

    If Check1(Index).Value = vbChecked Then
        OptionText(Index).BackColor = vbWhite
    Else
        OptionText(Index).BackColor = &H8000000B
    End If
      
End Sub

'--------------------------------------------------------
'   ｴﾝﾃﾞｨｱﾝｾﾚｸﾄ
'--------------------------------------------------------
Private Sub EndianSelect_Click(Index As Integer)

    If EndianSelect(0).Value = True Then
        Endian = BIG
    Else
        Endian = LITTLE
    End If

End Sub

'--------------------------------------------------------
'   ﾁｪｯｸﾎﾞﾀﾝｸﾘｯｸ後の動作
'--------------------------------------------------------
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim s_ReadModuleName(9) As Byte
    Dim s_ReadMacAdress(5) As Byte
    Dim SaveWorkFolder As String
      
    'ﾓｼﾞｭｰﾙ形名読出し
    If Index = 0 And Check1(0).Value = vbChecked Then
        OptionText(0).Text = ""
        '---FAM3と接続
        FAM3Connect
        '---ComMng_Moduleｸﾗｽﾌﾟﾛﾊﾟﾃｨ設定
        cmm.ModuleSlot = SelectedSlot
        cmm.EndianMode = Endian
        cmm.RomType = WordType
        '---ﾓｼﾞｭｰﾙ内CPUの動作をHALT
        Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
        Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
        '---ﾓｼﾞｭｰﾙ形名読出し
        Call cmm.ReadRom(RTOS_MODULENAME_ADDRESS, 10, s_ReadModuleName())
        For i = 0 To 8
            OptionText(0).Text = OptionText(0).Text + Chr(s_ReadModuleName(i))
        Next i
        '---FAM3から切断
        FAM3DisConnect
    End If
      
    'MACｱﾄﾞﾚｽ読出し
    If Index = 1 And Check1(1).Value = vbChecked Then
        OptionText(1).Text = ""
        '---FAM3と接続
        FAM3Connect
        '---ComMng_Moduleｸﾗｽﾌﾟﾛﾊﾟﾃｨ設定
        cmm.ModuleSlot = SelectedSlot
        cmm.EndianMode = Endian
        cmm.RomType = WordType
        '---ﾓｼﾞｭｰﾙ内CPUの動作をHALT
        Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
        Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
        '---MACｱﾄﾞﾚｽ読出し
        Call cmm.ReadRom(RTOS_MACADDRESS_ADDRESS, 6, s_ReadMacAdress())
        For i = 0 To 5
            OptionText(1).Text = OptionText(1).Text + Right("0" & Hex(s_ReadMacAdress(i)), 2)
        Next
        '---FAM3から切断
        FAM3DisConnect
    End If
      
    '基本ﾌﾞｰﾄﾌｧｲﾙを開く
    If Index = 2 And Check1(2).Value = vbChecked Then
        SaveWorkFolder = CurDir                           '現在のﾃﾞｨﾚｸﾄﾘ取得
        CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
        On Error GoTo Error_FirmFileOpen
            CommonDialog1.ShowOpen
        OptionText(2).Text = CommonDialog1.FileTitle      '基本ﾌﾞｰﾄﾌｧｲﾙ名
        SelectedFilePass = CommonDialog1.FileName         '基本ﾌﾞｰﾄﾌｧｲﾙのﾌﾙﾊﾟｽ
        ChDir SaveWorkFolder                              'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        'binﾌｧｲﾙのｻﾑ値表示
        FileSize = FileLen(SelectedFilePass)              '基本ﾌﾞｰﾄﾌｧｲﾙのｻｲｽﾞを取得
        FileNo = FreeFile(0)                              '使用可能なﾌｧｲﾙ番号を取得
        Open SelectedFilePass For Binary As FileNo        'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
        Get FileNo, FileSize - 1, SumValue()              '選択ﾌｧｲﾙのｻﾑ値を取得
        CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
        If CalSumValue > 65535 Then
            CalSumValue = CalSumValue - 65535
        End If
        FirmSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
        Close FileNo
    End If
      
    'OSﾌﾞｰﾄﾌｧｲﾙを開く
    If Index = 3 And Check1(3).Value = vbChecked Then
        SaveWorkFolder = CurDir                           '現在のﾃﾞｨﾚｸﾄﾘ取得
        CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
        On Error GoTo Error_SystemFileOpen
            CommonDialog1.ShowOpen
        OptionText(3).Text = CommonDialog1.FileTitle      'OSﾌﾞｰﾄﾌｧｲﾙ名
        SystemFilePass = CommonDialog1.FileName           'OSﾌﾞｰﾄﾌｧｲﾙのﾌﾙﾊﾟｽ
        ChDir SaveWorkFolder                              'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        'binﾌｧｲﾙのｻﾑ値表示
        SystemFileSize = FileLen(SystemFilePass)          'OSﾌﾞｰﾄﾌｧｲﾙのｻｲｽﾞを取得
        FileNo = FreeFile(0)                              '使用可能なﾌｧｲﾙ番号を取得
        Open SystemFilePass For Binary As FileNo          'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
        Get FileNo, SystemFileSize - 1, SumValue()        '選択ﾌｧｲﾙのｻﾑ値を取得
        CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
        If CalSumValue > 65535 Then
            CalSumValue = CalSumValue - 65535
        End If
        SystemSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
        Close FileNo
    End If

    Exit Sub

'ｷｬﾝｾﾙされた場合
Error_FirmFileOpen:
    If Err.Number = cdlCancel Then
        Check1(2).Value = False
        ChDir SaveWorkFolder    'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        Exit Sub
    End If
    Exit Sub
    
'ｷｬﾝｾﾙされた場合
Error_SystemFileOpen:
    If Err.Number = cdlCancel Then
        Check1(3).Value = False
        ChDir SaveWorkFolder    'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        Exit Sub
    End If
    
End Sub

'--------------------------------------------------------
'   ﾌｧｲﾙﾄﾞﾛｯﾌﾟ
'--------------------------------------------------------
Private Sub OptionText_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim yenpos As Integer
    Dim tmpFileSize As Long
    Dim tmpDropFilePass As String

    'ﾄﾞﾛｯﾌﾟされたものがﾌｧｲﾙか確認
    If Data.GetFormat(vbCFFiles) = False Then
        MsgBox "ﾌｧｲﾙ以外はﾄﾞﾛｯﾌﾟできません", vbCritical, "ERROR"
        Exit Sub
    End If
        
    'ﾄﾞﾛｯﾌﾟされたﾌｧｲﾙがbinﾌｧｲﾙか確認
    If InStr(LCase$(Data.Files(1)), ".bin") = 0 Then
        MsgBox "ﾊﾞｲﾅﾘﾌｧｲﾙ以外はﾄﾞﾛｯﾌﾟできません", vbCritical, "ERROR"
        Exit Sub
    End If
            
    '選択ﾌｧｲﾙのﾌﾙﾊﾟｽ
    tmpDropFilePass = Data.Files(1)
            
    '\を後ろから検索して文字のある場所を文字数で返す
    yenpos = InStrRev(tmpDropFilePass, "\")
    'ﾌｧｲﾙ名を取得して表示
    OptionText(Index).Text = Right$(tmpDropFilePass, Len(tmpDropFilePass) - yenpos)
            
    'binﾌｧｲﾙのｻﾑ値表示
    tmpFileSize = FileLen(tmpDropFilePass)                  '選択ﾌｧｲﾙのｻｲｽﾞを取得
    FileNo = FreeFile(0)                                    '使用可能なﾌｧｲﾙ番号を取得
    Open tmpDropFilePass For Binary As FileNo               'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
    Get FileNo, tmpFileSize - 1, SumValue()                 '選択ﾌｧｲﾙのｻﾑ値を取得
    CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
    If CalSumValue > 65535 Then
        CalSumValue = CalSumValue - 65535
    End If
    If Index = 2 Then
        FirmSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    ElseIf Index = 3 Then
        SystemSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    End If
    Close FileNo

    'ﾌｧｲﾙのﾊﾟｽとｻｲｽﾞを保存
    If Index = 2 Then
        SelectedFilePass = tmpDropFilePass
        FileSize = FileLen(SelectedFilePass)
    ElseIf Index = 3 Then
        SystemFilePass = tmpDropFilePass
        SystemFileSize = FileLen(SystemFilePass)
    End If

    'ﾁｪｯｸﾎﾞｯｸｽをﾁｪｯｸ
    Check1(Index).Value = vbChecked

End Sub

'--------------------------------------------------------
'   入力制限
'--------------------------------------------------------
Private Sub OptionText_KeyPress(Index As Integer, KeyAscii As Integer)

    If (Index = 2) Or (Index = 3) Then
        KeyAscii = 0
    End If

    '大文字入力のみ
    KeyCode = Chr(KeyAscii)
    KeyAscii = Asc(UCase(KeyCode))
    
End Sub

'--------------------------------------------------------
'   ﾌｫｰｶｽを表示しない
'--------------------------------------------------------
Private Sub OptionText_GotFocus(Index As Integer)

    If (Index = 2) Or (Index = 3) Then
        HideCaret (OptionText(Index).hwnd)  'ｷｬﾚｯﾄ非表示
    End If

End Sub


'--------------------------------------------------------
'
'   OKﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub OK_btn_Click()

    Dim i As Integer

    WriteMode = Normal_Write

    'ﾁｪｯｸﾎﾞｯｸｽが一つも選択されていない場合
    If (Check1(0).Value <> vbChecked) And (Check1(1).Value <> vbChecked) _
            And (Check1(2).Value <> vbChecked) And (Check1(3).Value <> vbChecked) Then
        MsgBox "書込みｱｲﾃﾑが一つも選択されていません", vbExclamation, "ERROR"
        Exit Sub
    End If

    '基本ﾌﾞｰﾄ書込み
    If Check1(2).Value = vbChecked Then
        WriteFirm = ENABLE
        WriteAddress = 0
        WriteSize = FileSize
    Else
        WriteFirm = DISABLE
    End If
    
    'OSﾌﾞｰﾄ書込み
    If Check1(3).Value = vbChecked Then
        WriteSystem = ENABLE
        SystemAddress = &H10000
    Else
        WriteSystem = DISABLE
    End If
    
    'ﾓｼﾞｭｰﾙ形名
    If Check1(0).Value = vbChecked Then
        If OptionText(0).Text = "" Then
            MsgBox "ﾓｼﾞｭｰﾙ形名を入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        If Len(OptionText(0).Text) <> 9 Then
            MsgBox "ﾓｼﾞｭｰﾙ形名を正しく入力してください（9桁）", vbExclamation, "ERROR"
            Exit Sub
        End If
        For i = 0 To 8
            WriteModuleNameData(i) = Asc(Mid(OptionText(0).Text, 1 + i, 1))
        Next i
        WriteModuleNameData(9) = &H20
        WriteName = ENABLE
    End If

    'MACｱﾄﾞﾚｽ
    If Check1(1).Value = vbChecked Then
        If OptionText(1).Text = "" Then
            MsgBox "MACｱﾄﾞﾚｽを入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        If Len(OptionText(1).Text) <> 12 Then
            MsgBox "MACｱﾄﾞﾚｽを正しく入力してください（12桁）", vbExclamation, "ERROR"
            Exit Sub
        End If
        For i = 0 To 5
            WriteMacData(i) = CLng("&H" + (Mid(OptionText(1).Text, 1 + i * 2, 2)))
        Next i
        WriteMAC = ENABLE
    End If
    
    Unload Me

    '残像対策
    Main_Form.Refresh
    
End Sub


'--------------------------------------------------------
'
'   ｷｬﾝｾﾙﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Cancel_btn_Click()

    WriteMode = Cancel_Write
    
    SelectedFilePass = PassForCancel
    FileSize = SizeForCancel
    
    Unload Me

End Sub

