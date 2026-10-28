VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form VerifySet_Form3 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Check"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4140
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Frame Frame1 
      Caption         =   "照合ｵﾌﾟｼｮﾝ"
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
      Begin VB.Frame Frame2 
         Caption         =   "ｴﾝﾃﾞｨｱﾝ"
         Height          =   615
         Left            =   360
         TabIndex        =   23
         Top             =   2040
         Width           =   2895
         Begin VB.OptionButton EndianSelect 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton EndianSelect 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label7 
            Caption         =   "Big"
            Height          =   255
            Left            =   840
            TabIndex        =   27
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "Little"
            Height          =   255
            Left            =   2160
            TabIndex        =   26
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   1
         Left            =   1560
         OLEDropMode     =   1  '手動
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   0
         Left            =   1560
         OLEDropMode     =   1  '手動
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   720
         Top             =   1440
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label FirmSum_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   2520
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   1800
         X2              =   1680
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   1680
         Y1              =   720
         Y2              =   840
      End
      Begin VB.Label Label10 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   760
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "基本ﾌﾞｰﾄ"
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Top             =   400
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "OSﾌﾞｰﾄ"
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   1600
         Width           =   615
      End
      Begin VB.Line Line3 
         X1              =   1680
         X2              =   1680
         Y1              =   1560
         Y2              =   1680
      End
      Begin VB.Line Line4 
         X1              =   1800
         X2              =   1680
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label SystemSum_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   2520
         TabIndex        =   15
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   525
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   825
      Width           =   135
   End
End
Attribute VB_Name = "VerifySet_Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*                                                          *
'*  VerifySet_Form3                                         *
'*                                                          *
'*      作成日  ：2007/06/18                                *
'*      作成者  ：澤原                                      *
'*      ｺﾒﾝﾄ    ：ROM照合 ｵﾌﾟｼｮﾝ画面（RTOS用）              *
'*                                                          *
'************************************************************

'Option Explicit

'---------------------------------------------
'外部API関数使用宣言
'---------------------------------------------
'---ｷｬﾚｯﾄを消す
Private Declare Function HideCaret Lib "user32.dll" (ByVal hwnd As Long) As Long
 

'--------------------------------------------------------
'
'   ROM照合ｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    Dim i As Integer

    '選択ﾌｧｲﾙのﾊﾟｽとｻｲｽﾞを保存しておく(ｷｬﾝｾﾙ時元に戻すため)
    PassForCancel = SelectedFilePass
    SizeForCancel = FileSize

    '表示初期化
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    SlotNo_lbl.Caption = SelectedSlot                                   'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名

    For i = 0 To 1
        OptionText(i).BackColor = &H8000000B
    Next
    
    If Main_Form.FileName_txt.Text <> "" Then
        If FileSize < 65536 Then    '基本ﾌﾞｰﾄ
            Check1(0).Value = vbChecked
            OptionText(0).Text = Main_Form.FileName_txt.Text        '書込みﾌｧｲﾙ名
            FirmSum_lbl.Caption = Main_Form.SumValue_lbl.Caption    '計算ｻﾑ値
            OptionText(0).BackColor = vbWhite
        Else                        'OSﾌﾞｰﾄ
            Check1(1).Value = vbChecked
            OptionText(1).Text = Main_Form.FileName_txt.Text        '書込みﾌｧｲﾙ名
            SystemSum_lbl.Caption = Main_Form.SumValue_lbl.Caption  '計算ｻﾑ値
            SystemFilePass = SelectedFilePass                       'OSﾌﾞｰﾄ用変数に代入
            SystemFileSize = FileSize                               'OSﾌﾞｰﾄ用変数に代入
            OptionText(1).BackColor = vbWhite
        End If
    End If
    
    EndianSelect(0).Value = True
    Endian = BIG
    
    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    VerifyAddress = 0
    VerifySize = 0
    MacWrittenAddress = 0
    
    'ROM照合用ﾌﾗｸﾞ初期化
    VerifyMode = Normal_Verify
    VerifyFirm = DISABLE
    VerifySystem = DISABLE
    
End Sub


'--------------------------------------------------------
'
'   照合ｵﾌﾟｼｮﾝ
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

    Dim SaveWorkFolder As String

    '基本ﾌﾞｰﾄﾌｧｲﾙを開く
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
        
    'OSﾌﾞｰﾄﾌｧｲﾙを開く
    ElseIf Index = 1 And Check1(1).Value = vbChecked Then
        SaveWorkFolder = CurDir                                 '現在のﾃﾞｨﾚｸﾄﾘ取得
        CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
        On Error GoTo Error_SystemFileOpen
            CommonDialog1.ShowOpen
        OptionText(1).Text = CommonDialog1.FileTitle            'OSﾌﾞｰﾄﾌｧｲﾙ名
        SystemFilePass = CommonDialog1.FileName                 'OSﾌﾞｰﾄﾌｧｲﾙのﾌﾙﾊﾟｽ
        ChDir SaveWorkFolder                                    'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        'binﾌｧｲﾙのｻﾑ値表示
        SystemFileSize = FileLen(SystemFilePass)                'OSﾌﾞｰﾄﾌｧｲﾙのｻｲｽﾞを取得
        FileNo = FreeFile(0)                                    '使用可能なﾌｧｲﾙ番号を取得
        Open SystemFilePass For Binary As FileNo                'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
        Get FileNo, SystemFileSize - 1, SumValue()              'OSﾌﾞｰﾄﾌｧｲﾙのｻﾑ値を取得
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
        Check1(0).Value = False
        ChDir SaveWorkFolder    'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
        Exit Sub
    End If
    Exit Sub

Error_SystemFileOpen:
    If Err.Number = cdlCancel Then
        Check1(1).Value = False
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
        SystemSum_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    End If
    Close FileNo

    'ﾁｪｯｸﾎﾞｯｸｽをﾁｪｯｸ
    If Index = 0 Then
        SelectedFilePass = tmpDropFilePass
        FileSize = FileLen(SelectedFilePass)
    ElseIf Index = 1 Then
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
    If (Check1(0).Value <> vbChecked) And (Check1(1).Value <> vbChecked) Then
        MsgBox "照合ｱｲﾃﾑが一つも選択されていません", vbExclamation, "ERROR"
        Exit Sub
    End If
    
    '基本ﾌﾞｰﾄ照合有り
    If Check1(0).Value = vbChecked Then
        VerifyFirm = ENABLE
        VerifyAddress = 0
        VerifySize = FileSize
    End If

    'OSﾌﾞｰﾄ照合有り
    If Check1(1).Value = vbChecked Then
        VerifySystem = ENABLE
        SystemAddress = &H10000
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

    VerifyMode = Cancel_Verify
    
    SelectedFilePass = PassForCancel
    FileSize = SizeForCancel
    
    Unload Me

End Sub

