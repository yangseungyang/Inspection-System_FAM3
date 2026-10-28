VERSION 5.00
Begin VB.Form VerifySet_Form1 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Check"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3990
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.Frame Frame1 
      Caption         =   "照合ｵﾌﾟｼｮﾝ"
      Height          =   1215
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Width           =   3495
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   20
         Top             =   700
         Width           =   1335
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   19
         Top             =   340
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "照合ｻｲｽﾞ(KByte)"
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "照合ｱﾄﾞﾚｽ(Hex)"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   380
         Width           =   1455
      End
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   1930
      TabIndex        =   17
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   1930
      TabIndex        =   16
      Top             =   525
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   1930
      TabIndex        =   12
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   840
      TabIndex        =   11
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "："
      Height          =   255
      Left            =   1930
      TabIndex        =   7
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label SumValue_lbl 
      Caption         =   "SumValue"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1425
      Width           =   1335
   End
   Begin VB.Label CheckFile_lbl 
      Caption         =   "WriteFile"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   1125
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "："
      Height          =   255
      Left            =   1930
      TabIndex        =   4
      Top             =   1125
      Width           =   135
   End
   Begin VB.Label Label15 
      Caption         =   "ｻﾑ値"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1425
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "照合ﾌｧｲﾙ"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1125
      Width           =   975
   End
End
Attribute VB_Name = "VerifySet_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  VerifySet_Form1                                     *
'*                                                      *
'*      作成日  ：2007/05/09                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM照合 ｵﾌﾟｼｮﾝ画面（ﾉｰﾏﾙ）            *
'*                                                      *
'********************************************************


'--------------------------------------------------------
'
'   ROM照合ｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    '表示初期化
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    SlotNo_lbl.Caption = SelectedSlot                                   'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名
    CheckFile_lbl.Caption = Main_Form.FileName_txt.Text                 '照合ﾌｧｲﾙ名
    SumValue_lbl.Caption = Main_Form.SumValue_lbl.Caption               '計算ｻﾑ値

    OptionText(0).Text = 0
    OptionText(1).Text = FileLen(SelectedFilePass) / 1024
    
    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    VerifyAddress = 0
    VerifySize = 0
    Endian = BIG
    
    'ROM照合用ﾌﾗｸﾞ初期化
    VerifyMode = Normal_Verify
    VerifyFirm = ENABLE
    
End Sub


'--------------------------------------------------------
'
'   照合ｵﾌﾟｼｮﾝ
'
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
    If CLng("&H" + OptionText(0).Text) Mod 512 <> 0 Then
        MsgBox "照合ｱﾄﾞﾚｽは512Byte単位($200)で設定してください", vbExclamation, "ERROR"
        Exit Sub
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
  
    'CMDGA特別ｴﾗｰﾁｪｯｸ
    If ModuleType = CMDGA_TYPE Then
        If VerifyAddress + VerifySize > 131072 Then
            MsgBox "CMDGAﾀｲﾌﾟのﾓｼﾞｭｰﾙはFlashROMの$20000番地以降へｱｸｾｽできません", vbExclamation, "ERROR"
            Exit Sub
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
