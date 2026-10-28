VERSION 5.00
Begin VB.Form CopySet_Form2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Copy"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   4710
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.ComboBox TargetSlot_lst 
      Height          =   300
      Left            =   2880
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "ｺﾋﾟｰｵﾌﾟｼｮﾝ"
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   4215
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   2
         Left            =   2520
         MaxLength       =   12
         TabIndex        =   20
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   1
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   0
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   18
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "MACｱﾄﾞﾚｽ"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   1560
         Width           =   1800
      End
      Begin VB.Line Line1 
         X1              =   730
         X2              =   600
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line5 
         X1              =   730
         X2              =   600
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label11 
         Caption         =   "MACｱﾄﾞﾚｽ書換え"
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   840
         Width           =   1815
      End
      Begin VB.Line Line6 
         X1              =   615
         X2              =   615
         Y1              =   1080
         Y2              =   1680
      End
      Begin VB.Label Label12 
         Caption         =   "MAC先頭番地(Hex)"
         Height          =   255
         Left            =   855
         TabIndex        =   15
         Top             =   1245
         Width           =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "ROMｻｲｽﾞ（Mbit）"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.ComboBox SourceSlot_lst 
      Height          =   300
      Left            =   1200
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   0
      Top             =   1560
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2160
      Picture         =   "CopySet_Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2310
      TabIndex        =   13
      Top             =   585
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   585
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2310
      TabIndex        =   10
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Caption         =   "ｺﾋﾟｰ元ｽﾛｯﾄ番号"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "ｺﾋﾟｰ先ｽﾛｯﾄ番号"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "CopySet_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*                                                          *
'*  CopySet_Form1                                           *
'*                                                          *
'*      作成日  ：2007/06/21                                *
'*      作成者  ：澤原                                      *
'*      ｺﾒﾝﾄ    ：ROMｺﾋﾟｰ ｵﾌﾟｼｮﾝ画面（Ethernet用）          *
'*                                                          *
'************************************************************


'--------------------------------------------------------
'
'   ROMｺﾋﾟｰｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    Dim i As Integer
    
    '表示初期化
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名

    SourceSlot_lst.AddItem SelectedSlot             'ｺﾋﾟｰ元ｽﾛｯﾄ初期化
    SourceSlot_lst.Text = SourceSlot_lst.List(0)
    For i = 2 To 16
        If i <> SelectedSlot Then
            TargetSlot_lst.AddItem i
        End If
    Next i
    TargetSlot_lst.ListIndex = 0                    'ｺﾋﾟｰ先ｽﾛｯﾄ初期値

    OptionText(0).Text = ""                         'ROMｻｲｽﾞ初期値
    OptionText(1).Text = "4000"                     'MAC先頭番地初期値
    OptionText(2).Text = ""                         'MACｱﾄﾞﾚｽ初期値
    For i = 1 To 2
        OptionText(i).Enabled = False
        OptionText(i).BackColor = &H8000000B
    Next

    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    CopyAddress = 0
    CopySize = 0
    
    'ROMｺﾋﾟｰ用ﾌﾗｸﾞ初期化
    CopyMode = Normal_Copy
    ChangeMAC = DISABLE
    
End Sub


'--------------------------------------------------------
'
'   ｺﾋﾟｰｵﾌﾟｼｮﾝ
'
'--------------------------------------------------------
'--------------------------------------------------------
'   MACｱﾄﾞﾚｽ書換え
'--------------------------------------------------------
Private Sub Check1_Click()

    If Check1.Value = vbChecked Then
        OptionText(1).Enabled = True
        OptionText(1).BackColor = vbWhite
        OptionText(2).Enabled = True
        OptionText(2).BackColor = vbWhite
    Else
        OptionText(1).Enabled = False
        OptionText(1).BackColor = &H8000000B
        OptionText(2).Enabled = False
        OptionText(2).BackColor = &H8000000B
    End If

End Sub

'--------------------------------------------------------
'   ﾃｷｽﾄﾎﾞｯｸｽ入力制限
'--------------------------------------------------------
Private Sub OptionText_KeyPress(Index As Integer, KeyAscii As Integer)

    '大文字入力のみ
    KeyCode = Chr(KeyAscii)
    KeyAscii = Asc(UCase(KeyCode))
    'BackSpaceとDelete有効
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    'ｷｰ入力制限
    If Index = 0 Then   '数値入力
        If KeyAscii < Asc("0") Or Asc("9") < KeyAscii Then
            KeyAscii = 0
        End If
    Else                'ｱﾄﾞﾚｽ入力
        If KeyAscii < Asc("0") Or Asc("F") < KeyAscii Then
            KeyAscii = 0
        End If
    End If
    
End Sub


'--------------------------------------------------------
'
'   OKﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub OK_btn_Click()

    Dim i As Integer

    'ROMｻｲｽﾞ
    If OptionText(0).Text = "" Then
        MsgBox "ROMｻｲｽﾞを入力してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If CLng(OptionText(0).Text) <= 0 Then
        MsgBox "ROMｻｲｽﾞが範囲外です", vbExclamation, "ERROR"
        Exit Sub
    End If
    
    'MACｱﾄﾞﾚｽ
    If Check1.Value = vbChecked Then
        If OptionText(1).Text = "" Then
            MsgBox "MAC先頭番地を入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        If OptionText(2).Text = "" Then
            MsgBox "MACｱﾄﾞﾚｽを入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        If Len(OptionText(2).Text) <> 12 Then
            MsgBox "MACｱﾄﾞﾚｽを正しく入力してください（12桁）", vbExclamation, "ERROR"
            Exit Sub
        End If
        'MAC先頭番地取得
        MacWriteAddress = CLng("&H" + OptionText(1).Text)
        If MacWriteAddress < 0 Then
            MacWriteAddress = MacWriteAddress + &H10000
        End If
        'MACｱﾄﾞﾚｽ取得
        For i = 0 To 5
            WriteMacData(i) = CLng("&H" + (Mid(OptionText(2).Text, 1 + i * 2, 2)))
        Next i
        WriteMacData(6) = 0
        WriteMacData(7) = Not WriteMacData(3)
        WriteMacData(8) = Not WriteMacData(4)
        WriteMacData(9) = Not WriteMacData(5)
        ChangeMAC = ENABLE
    End If

    'ｵﾌﾟｼｮﾝ情報をｾｯﾄ
    CopySize = CLng(OptionText(0).Text)
    TargetSlot = CLng(TargetSlot_lst.Text)
    
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

    CopyMode = Cancel_Copy
    Unload Me
    
End Sub
