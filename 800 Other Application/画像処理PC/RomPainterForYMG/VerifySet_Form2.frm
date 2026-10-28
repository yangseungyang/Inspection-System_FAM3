VERSION 5.00
Begin VB.Form VerifySet_Form2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Check"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4635
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   960
      TabIndex        =   23
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "照合ｵﾌﾟｼｮﾝ"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   4095
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   2
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   25
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "MAC先頭番地(Hex)"
         Height          =   255
         Left            =   855
         TabIndex        =   21
         Top             =   1605
         Width           =   1800
      End
      Begin VB.Line Line6 
         X1              =   615
         X2              =   615
         Y1              =   1440
         Y2              =   1680
      End
      Begin VB.Line Line5 
         X1              =   730
         X2              =   600
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label11 
         Caption         =   "MACｱﾄﾞﾚｽ非照合"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "照合ｱﾄﾞﾚｽ(Hex)"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   380
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "照合ｻｲｽﾞ(KByte)"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   740
         Width           =   1695
      End
   End
   Begin VB.Label Label16 
      Caption         =   "照合ﾌｧｲﾙ"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   1125
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "ｻﾑ値"
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   1425
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   1125
      Width           =   135
   End
   Begin VB.Label CheckFile_lbl 
      Caption         =   "WriteFile"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   1125
      Width           =   1815
   End
   Begin VB.Label SumValue_lbl 
      Caption         =   "SumValue"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   1425
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   240
      Width           =   135
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   525
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   825
      Width           =   135
   End
End
Attribute VB_Name = "VerifySet_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*                                                          *
'*  VerifySet_Form2                                         *
'*                                                          *
'*      作成日  ：2007/05/09                                *
'*      作成者  ：澤原                                      *
'*      ｺﾒﾝﾄ    ：ROM照合 ｵﾌﾟｼｮﾝ画面（Ethernet,FL-net用）   *
'*                                                          *
'************************************************************


'--------------------------------------------------------
'
'   ROM照合ｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    Dim i As Integer

    '表示初期化
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    SlotNo_lbl.Caption = SelectedSlot                                   'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名
    CheckFile_lbl.Caption = Main_Form.FileName_txt.Text                 '照合ﾌｧｲﾙ名
    SumValue_lbl.Caption = Main_Form.SumValue_lbl.Caption               '計算ｻﾑ値

    OptionText(0).Text = 0
    OptionText(1).Text = FileLen(SelectedFilePass) / 1024
    OptionText(2).Enabled = False
    OptionText(2).BackColor = &H8000000B
    OptionText(2).Text = "4000"
    
    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    VerifyAddress = 0
    VerifySize = 0
    MacWrittenAddress = 0
    Endian = BIG
    
    'ROM照合用ﾌﾗｸﾞ初期化
    VerifyMode = Normal_Verify
    VerifyFirm = ENABLE
    VerifyMAC = ENABLE
    
End Sub


'--------------------------------------------------------
'
'   照合ｵﾌﾟｼｮﾝ
'
'--------------------------------------------------------
'--------------------------------------------------------
'   MACｱﾄﾞﾚｽ非照合ﾁｪｯｸﾎﾞｯｸｽ
'--------------------------------------------------------
Private Sub Check1_Click()

    If Check1.Value = vbChecked Then
        OptionText(2).Enabled = True
        OptionText(2).BackColor = vbWhite
    Else
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
    
    '照合ｱﾄﾞﾚｽ
    If OptionText(0).Text = "" Then
        MsgBox "照合ｱﾄﾞﾚｽを入力してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If CLng("&H" + OptionText(0).Text) Mod 2 <> 0 Then
        MsgBox "照合ｱﾄﾞﾚｽは偶数番地を指定してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If ModuleType = FLNET_TYPE Then
        If CLng("&H" + OptionText(0).Text) Mod 8192 <> 0 Then
            MsgBox "照合ｱﾄﾞﾚｽは8KByte単位($2000)で設定してください", vbExclamation, "ERROR"
            Exit Sub
        End If
    Else
        If CLng("&H" + OptionText(0).Text) Mod 512 <> 0 Then
            MsgBox "照合ｱﾄﾞﾚｽは512Byte単位($200)で設定してください", vbExclamation, "ERROR"
            Exit Sub
        End If
    End If
    '照合ｻｲｽﾞ
    If OptionText(1).Text = "" Then
        MsgBox "照合ｻｲｽﾞを入力してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If CLng(OptionText(1).Text) <= 0 Then
        MsgBox "照合ｻｲｽﾞが範囲外です", vbExclamation, "ERROR"
        Exit Sub
    End If
    
    'ｵﾌﾟｼｮﾝ情報をｾｯﾄ
    VerifyAddress = CLng("&H" + OptionText(0).Text)
    If VerifyAddress < 0 Then
        VerifyAddress = VerifyAddress + 65536
    End If
    VerifySize = CLng(OptionText(1).Text) * 1024
    
    '照合範囲ﾁｪｯｸ
    If VerifyAddress + VerifySize > FileSize Then
        MsgBox "照合範囲が比較ﾌｧｲﾙのｻｲｽﾞを超えています", vbExclamation, "ERROR"
        Exit Sub
    End If
    
    'MACｱﾄﾞﾚｽ非照合
    If Check1.Value = vbChecked Then
        If OptionText(2).Text = "" Then
            MsgBox "MACｱﾄﾞﾚｽが書き込まれている先頭番地を入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        If CLng("&H" + OptionText(2).Text) Mod 2 <> 0 Then
            MsgBox "MAC先頭番地は偶数番地を指定してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        VerifyMAC = DISABLE
        MacWrittenAddress = CLng("&H" + OptionText(2).Text)
        If MacWrittenAddress < 0 Then
            MacWrittenAddress = MacWrittenAddress + 65536
        End If
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
    Unload Me

End Sub

