VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form WriteSet_Form4 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Write"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4695
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "書込みｵﾌﾟｼｮﾝ"
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   27
         Top             =   390
         Width           =   255
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   2
         Left            =   2160
         OLEDropMode     =   1  '手動
         TabIndex        =   22
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   1
         Left            =   2160
         OLEDropMode     =   1  '手動
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   240
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   0
         Left            =   2160
         OLEDropMode     =   1  '手動
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   1100
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   750
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   2280
         X2              =   2280
         Y1              =   720
         Y2              =   840
      End
      Begin VB.Line Line1 
         X1              =   2400
         X2              =   2280
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label FirmSum_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   3120
         TabIndex        =   25
         Top             =   720
         Width           =   735
      End
      Begin VB.Label CfgInfoSum_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   3120
         TabIndex        =   24
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label DevInfoSum_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   3120
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.Line Line6 
         X1              =   2400
         X2              =   2280
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line5 
         X1              =   2280
         X2              =   2280
         Y1              =   2160
         Y2              =   2280
      End
      Begin VB.Label Label13 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   2190
         Width           =   615
      End
      Begin VB.Line Line4 
         X1              =   2400
         X2              =   2280
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line3 
         X1              =   2280
         X2              =   2280
         Y1              =   1440
         Y2              =   1560
      End
      Begin VB.Label Label12 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   1470
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "ﾃﾞﾊﾞｲｽﾈｯﾄ情報"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "ﾌｧｰﾑｳｪｱ"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   390
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   2300
      TabIndex        =   20
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2300
      TabIndex        =   16
      Top             =   240
      Width           =   135
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2300
      TabIndex        =   12
      Top             =   525
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   825
      Width           =   135
   End
End
Attribute VB_Name = "WriteSet_Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  WriteSet_Form4                                      *
'*                                                      *
'*      作成日  ：2007/06/08                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM書込み ｵﾌﾟｼｮﾝ画面（DeviceNet用）   *
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

    '選択ﾌｧｲﾙのﾊﾟｽとｻｲｽﾞを保存しておく(ｷｬﾝｾﾙ時元に戻すため)
    PassForCancel = SelectedFilePass
    SizeForCancel = FileSize
    
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    SlotNo_lbl.Caption = SelectedSlot                                   'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名

    If Main_Form.FileName_txt.Text <> "" Then
        Check1(0).Value = vbChecked
        OptionText(0).Text = Main_Form.FileName_txt.Text
        FirmSum_lbl.Caption = Main_Form.SumValue_lbl.Caption
    Else
        OptionText(0).BackColor = &H8000000B
    End If
    
    For i = 1 To 2
        OptionText(i).BackColor = &H8000000B
        OptionText(i).Text = ""
    Next

    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    WriteAddress = 0
    WriteSize = 0
    Endian = BIG
    SaveFilePass = SelectedFilePass         '元のﾌｧｲﾙのﾊﾟｽを保存しておく

    'ROM書込み用ﾌﾗｸﾞ初期化
    WriteMode = Normal_Write
    WriteFirm = DISABLE
    WriteDevInfo = DISABLE
    WriteCfgInfo = DISABLE

End Sub


'--------------------------------------------------------
'
'   書込みｵﾌﾟｼｮﾝ
'
'--------------------------------------------------------
'--------------------------------------------------------
'   ﾁｪｯｸﾎﾞｯｸｽ
'--------------------------------------------------------
Private Sub Check1_Click(Index As Integer)

    If Check1(Index).Value = vbChecked Then
        OptionText(Index).BackColor = vbWhite
    Else
        OptionText(Index).BackColor = &H8000000B
    End If

End Sub

'--------------------------------------------------------
'   ﾌｧｲﾙ選択
'--------------------------------------------------------
Private Sub Check1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim SaveWorkFolder As String
    
    'F/Wﾌｧｲﾙを開く
    If Index = 0 And Check1(0).Value = vbChecked Then
        SaveWorkFolder = CurDir                                 '現在のﾃﾞｨﾚｸﾄﾘ取得
        CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
        On Error GoTo Error_FirmFileOpen
            CommonDialog1.ShowOpen
        OptionText(0).Text = CommonDialog1.FileTitle            'F/Wﾌｧｲﾙ名
        SelectedFilePass = CommonDialog1.FileName               'F/Wﾌｧｲﾙのﾌﾙﾊﾟｽ
        ChDir SaveWorkFolder                                    'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        'binﾌｧｲﾙのｻﾑ値表示
        FileSize = FileLen(SelectedFilePass)                    'F/Wﾌｧｲﾙのｻｲｽﾞを取得
        FileNo = FreeFile(0)                                    '使用可能なﾌｧｲﾙ番号を取得
        Open SelectedFilePass For Binary As FileNo              'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
        Get FileNo, FileSize - 1, SumValue()                    'ｻﾑ値取得
        CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
        If CalSumValue > 65535 Then
            CalSumValue = CalSumValue - 65535
        End If
        FirmSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
        Close FileNo                                            'ﾊﾞｲﾅﾘﾌｧｲﾙを閉じる
    
    'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙを開く
    ElseIf Index = 1 And Check1(1).Value = vbChecked Then
        SaveWorkFolder = CurDir                          '現在のﾃﾞｨﾚｸﾄﾘ取得
        CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
        On Error GoTo Error_DevInfoFileOpen
            CommonDialog1.ShowOpen
        OptionText(1).Text = CommonDialog1.FileTitle     'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙ名
        DevInfoFilePass = CommonDialog1.FileName         'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙのﾌﾙﾊﾟｽ
        ChDir SaveWorkFolder                             'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        'binﾌｧｲﾙのｻﾑ値表示
        DevInfoFileSize = FileLen(DevInfoFilePass)       'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙのｻｲｽﾞを取得
        FileNo = FreeFile(0)                             '使用可能なﾌｧｲﾙ番号を取得
        Open DevInfoFilePass For Binary As FileNo        'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
        Get FileNo, DevInfoFileSize - 1, SumValue()      '選択ﾌｧｲﾙのｻﾑ値を取得
        CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
        If CalSumValue > 65535 Then
            CalSumValue = CalSumValue - 65535
        End If
        DevInfoSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
        Close FileNo                                     'ﾊﾞｲﾅﾘﾌｧｲﾙを閉じる
    
    'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙを開く
    ElseIf Index = 2 And Check1(2).Value = vbChecked Then
        SaveWorkFolder = CurDir                          '現在のﾃﾞｨﾚｸﾄﾘ取得
        CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
        On Error GoTo Error_CfgInfoFileOpen
            CommonDialog1.ShowOpen
        OptionText(2).Text = CommonDialog1.FileTitle     'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙ名
        CfgInfoFilePass = CommonDialog1.FileName         'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙのﾌﾙﾊﾟｽ
        ChDir SaveWorkFolder                             'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        'binﾌｧｲﾙのｻﾑ値表示
        CfgInfoFileSize = FileLen(CfgInfoFilePass)       'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙのｻｲｽﾞを取得
        FileNo = FreeFile(0)                             '使用可能なﾌｧｲﾙ番号を取得
        Open CfgInfoFilePass For Binary As FileNo        'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
        Get FileNo, CfgInfoFileSize - 1, SumValue()      '選択ﾌｧｲﾙのｻﾑ値を取得
        CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
        If CalSumValue > 65535 Then
            CalSumValue = CalSumValue - 65535
        End If
        CfgInfoSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
        Close FileNo                                     'ﾊﾞｲﾅﾘﾌｧｲﾙを閉じる
    End If

    Exit Sub
    
'ｷｬﾝｾﾙされた場合
Error_FirmFileOpen:
    If Err.Number = cdlCancel Then
        Check1(0).Value = False
        ChDir SaveWorkFolder    'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        Exit Sub
    End If
    Exit Sub
    
Error_DevInfoFileOpen:
    If Err.Number = cdlCancel Then
        Check1(1).Value = False
        ChDir SaveWorkFolder    'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        Exit Sub
    End If
    Exit Sub

Error_CfgInfoFileOpen:
    If Err.Number = cdlCancel Then
        Check1(2).Value = False
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
    If Index = 0 Then
        FirmSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    ElseIf Index = 1 Then
        DevInfoSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    ElseIf Index = 2 Then
        CfgInfoSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    End If
    Close FileNo

    'ﾌｧｲﾙのﾊﾟｽとｻｲｽﾞを保存
    If Index = 0 Then
        SelectedFilePass = tmpDropFilePass
        FileSize = FileLen(SelectedFilePass)
    ElseIf Index = 1 Then
        DevInfoFilePass = tmpDropFilePass
        DevInfoFileSize = FileLen(DevInfoFilePass)
    ElseIf Index = 2 Then
        CfgInfoFilePass = tmpDropFilePass
        CfgInfoFileSize = FileLen(CfgInfoFilePass)
    End If

    'ﾁｪｯｸﾎﾞｯｸｽをﾁｪｯｸ
    Check1(Index).Value = vbChecked

End Sub

'--------------------------------------------------------
'   入力制限
'--------------------------------------------------------
Private Sub OptionText_KeyPress(Index As Integer, KeyAscii As Integer)

    KeyAscii = 0

End Sub

'--------------------------------------------------------
'   ﾌｫｰｶｽを表示しない
'--------------------------------------------------------
Private Sub OptionText_GotFocus(Index As Integer)

    HideCaret (OptionText(Index).hwnd)  'ｷｬﾚｯﾄ非表示

End Sub


'--------------------------------------------------------
'
'   OKﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub OK_btn_Click()
    
    'ﾁｪｯｸﾎﾞｯｸｽが一つも選択されていない場合
    If (Check1(0).Value <> vbChecked) And (Check1(1).Value <> vbChecked) And (Check1(2).Value <> vbChecked) Then
        MsgBox "書込みｱｲﾃﾑが一つも選択されていません", vbExclamation, "ERROR"
        Exit Sub
    End If
    
    'F/W書込み有り
    If Check1(0).Value = vbChecked Then
        WriteFirm = ENABLE
        WriteAddress = 0
        WriteSize = FileSize
    End If

    'ﾃﾞﾊﾞｲｽﾈｯﾄ情報の書込み有り
    If Check1(1).Value = vbChecked Then
        WriteDevInfo = ENABLE
        DevInfoAddress = &H20000
    End If

    'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報の書込み有り
    If Check1(2).Value = vbChecked Then
        WriteCfgInfo = ENABLE
        CfgInfoAddress = &H30000
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


