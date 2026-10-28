VERSION 5.00
Begin VB.Form EraseSet_Form 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Erase"
   ClientHeight    =   4065
   ClientLeft      =   4650
   ClientTop       =   3720
   ClientWidth     =   4470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4470
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2400
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "消去ｵﾌﾟｼｮﾝ"
      Height          =   1935
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   3975
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   16
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton EraseMode_btn 
         Caption         =   "ｾｸﾀ消去"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton EraseMode_btn 
         Caption         =   "全消去"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Line Line3 
         X1              =   720
         X2              =   960
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         X1              =   720
         X2              =   960
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         X1              =   720
         X2              =   720
         Y1              =   960
         Y2              =   1560
      End
      Begin VB.Label Label6 
         Caption         =   "消去ｻｲｽﾞ(KByte)"
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "ｾｸﾀｱﾄﾞﾚｽ(Hex)"
         Height          =   225
         Left            =   1080
         TabIndex        =   9
         Top             =   1125
         Width           =   1170
      End
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2180
      TabIndex        =   13
      Top             =   820
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2180
      TabIndex        =   12
      Top             =   530
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   820
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   530
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2180
      TabIndex        =   3
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   820
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   530
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "EraseSet_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  EraseSet_Form                                       *
'*                                                      *
'*      作成日  ：2007/02/15                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM消去 ｵﾌﾟｼｮﾝ画面                    *
'*                                                      *
'********************************************************


'--------------------------------------------------------
'
'   ROM消去ｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    Dim i As Integer

    '表示初期化
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text
    SlotNo_lbl.Caption = SelectedSlot
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)
    For i = 0 To 1
        OptionText(i).Text = ""
        OptionText(i).BackColor = &H8000000B
    Next
    EraseMode_btn(0).Enabled = True     'ﾃﾞﾌｫﾙﾄは全消去を選択

    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    EraseAddress = 0
    EraseSize = 0

End Sub


'--------------------------------------------------------
'
'   消去ｵﾌﾟｼｮﾝ
'
'--------------------------------------------------------
'--------------------------------------------------------
'   消去ﾓｰﾄﾞ選択
'--------------------------------------------------------
Private Sub EraseMode_btn_Click(Index As Integer)

    If Index = 0 Then
        '全消去選択
        OptionText(0).Enabled = False
        OptionText(0).BackColor = &H8000000B
        OptionText(1).Enabled = False
        OptionText(1).BackColor = &H8000000B
    Else
        'ｾｸﾀ消去選択
        OptionText(0).Enabled = True
        OptionText(0).BackColor = &H80000009
        OptionText(1).Enabled = True
        OptionText(1).BackColor = &H80000009
        OptionText(0).SetFocus
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
    If Index = 1 Then   '数値入力
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

    If EraseMode_btn(0) = True Then
        EraseMode = Tip_Erase
        EraseAddress = 0
    Else
        EraseMode = Sector_Erase
        If OptionText(0).Text = "" Then
            MsgBox "消去ｱﾄﾞﾚｽを入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        
        If OptionText(1).Text = "" Then
            MsgBox "消去ｻｲｽﾞを入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        If CLng(OptionText(1).Text) <= 0 Then
            MsgBox "消去ｻｲｽﾞが範囲外です", vbExclamation, "ERROR"
            Exit Sub
        End If
        
        EraseAddress = CLng("&H" + OptionText(0).Text)
        If EraseAddress < 0 Then
            EraseAddress = EraseAddress + &H10000
        End If
        EraseSize = CLng(OptionText(1).Text)
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
    
    EraseMode = Cancel_Erase

    Unload Me

End Sub

