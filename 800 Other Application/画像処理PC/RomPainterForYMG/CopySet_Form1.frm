VERSION 5.00
Begin VB.Form CopySet_Form1 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Copy"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   3990
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.ComboBox SourceSlot_lst 
      Height          =   300
      Left            =   840
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   12
      Top             =   1560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "ｺﾋﾟｰｵﾌﾟｼｮﾝ"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   3495
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   0
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   14
         Top             =   340
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "ROMｻｲｽﾞ(MBit)"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   390
         Width           =   1335
      End
   End
   Begin VB.ComboBox TargetSlot_lst 
      Height          =   300
      Left            =   2520
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1800
      Picture         =   "CopySet_Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Caption         =   "ｺﾋﾟｰ先ｽﾛｯﾄ番号"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Caption         =   "ｺﾋﾟｰ元ｽﾛｯﾄ番号"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   1950
      TabIndex        =   3
      Top             =   240
      Width           =   135
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   585
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   1950
      TabIndex        =   0
      Top             =   585
      Width           =   135
   End
End
Attribute VB_Name = "CopySet_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  CopySet_Form1                                       *
'*                                                      *
'*      作成日  ：2007/06/21                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROMｺﾋﾟｰ ｵﾌﾟｼｮﾝ画面(ﾉｰﾏﾙ)              *
'*                                                      *
'********************************************************


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
    
    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    CopyAddress = 0
    CopySize = 0
    
    'ROMｺﾋﾟｰ用ﾌﾗｸﾞ初期化
    CopyMode = Normal_Copy
    ChangeMAC = DISABLE
    
End Sub



'--------------------------------------------------------
'
'   ｺﾋﾟｰｵﾌﾟｼｮﾝ （ﾃｷｽﾄﾎﾞｯｸｽ入力制限）
'
'--------------------------------------------------------
Private Sub OptionText_KeyPress(Index As Integer, KeyAscii As Integer)

    '大文字入力のみ
    KeyCode = Chr(KeyAscii)
    KeyAscii = Asc(UCase(KeyCode))
    'BackSpaceとDelete有効
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    '数値入力のみ
    If KeyAscii < Asc("0") Or Asc("9") < KeyAscii Then
        KeyAscii = 0
    End If
    
End Sub


'--------------------------------------------------------
'
'   OKﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub OK_btn_Click()

    'ROMｻｲｽﾞ
    If OptionText(0).Text = "" Then
        MsgBox "ROMｻｲｽﾞを入力してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If CLng(OptionText(0).Text) <= 0 Then
        MsgBox "ROMｻｲｽﾞが範囲外です", vbExclamation, "ERROR"
        Exit Sub
    End If
    If ModuleType = CMDGA_TYPE Then
        If CLng(OptionText(0).Text) <> 1 Then
            MsgBox "ｺﾏﾝﾄﾞｹﾞｰﾄｱﾚｲ使用ﾓｼﾞｭｰﾙは1Mbit限定でｺﾋﾟｰして下さい", vbExclamation, "ERROR"
            Exit Sub
        End If
    End If

    'ｵﾌﾟｼｮﾝ情報をｾｯﾄ
    TargetSlot = CLng(TargetSlot_lst.Text)
    CopySize = CLng(OptionText(0).Text)

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

