VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ReadSet_Form2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Read"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4665
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "読出しｵﾌﾟｼｮﾝ"
      Height          =   2175
      Left            =   240
      TabIndex        =   10
      Top             =   1320
      Width           =   4095
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   2
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   20
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   19
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   255
      End
      Begin VB.Line Line5 
         X1              =   730
         X2              =   600
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line6 
         X1              =   615
         X2              =   615
         Y1              =   1440
         Y2              =   1680
      End
      Begin VB.Label Label12 
         Caption         =   "MAC先頭番地(Hex)"
         Height          =   255
         Left            =   855
         TabIndex        =   17
         Top             =   1605
         Width           =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "MACｱﾄﾞﾚｽ初期化"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   1230
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "読出しｻｲｽﾞ(KByte)"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "読出しｱﾄﾞﾚｽ(Hex)"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   390
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2290
      TabIndex        =   9
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2290
      TabIndex        =   8
      Top             =   525
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2290
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   825
      Width           =   135
   End
End
Attribute VB_Name = "ReadSet_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*                                                          *
'*  ReadSet_Form                                            *
'*                                                          *
'*      作成日  ：2007/04/03                                *
'*      作成者  ：澤原                                      *
'*      ｺﾒﾝﾄ    ：ROM読出し ｵﾌﾟｼｮﾝ画面(Ethernet,FL-net系)   *
'*                                                          *
'************************************************************


'--------------------------------------------------------
'
'   ROM読出しｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    SlotNo_lbl.Caption = SelectedSlot                                   'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名

    OptionText(0).Text = "0"
    OptionText(2).Enabled = False
    OptionText(2).BackColor = &H8000000B
    OptionText(2).Text = "4000"
    
    'ROM読出し用ﾌﾗｸﾞ初期化
    ReadMode = Normal_Read
    InitMacAddress = DISABLE

End Sub


'--------------------------------------------------------
'
'   MACｱﾄﾞﾚｽ初期化の有無
'
'--------------------------------------------------------
Private Sub Check1_Click()
    
    If Check1.Value = vbChecked Then
        InitMacAddress = ENABLE
        OptionText(2).Enabled = True
        OptionText(2).BackColor = vbWhite
    Else
        InitMacAddress = DISABLE
        OptionText(2).Enabled = False
        OptionText(2).BackColor = &H8000000B
    End If

End Sub

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

    Dim res As Integer
    Dim SaveWorkFolder As String
    Dim tmpfso As New FileSystemObject

    '読出し開始ｱﾄﾞﾚｽとｻｲｽﾞ
    If OptionText(0).Text = "" Then
        MsgBox "読出し開始ｱﾄﾞﾚｽを入力してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If ModuleType = FLNET_TYPE Then
        If CLng("&H" + OptionText(0).Text) Mod 8192 <> 0 Then
            MsgBox "読出し開始ｱﾄﾞﾚｽは8KByte単位($2000)で設定してください", vbExclamation, "ERROR"
            Exit Sub
        End If
    Else
        If CLng("&H" + OptionText(0).Text) Mod 512 <> 0 Then
            MsgBox "読出し開始ｱﾄﾞﾚｽは512Byte単位($200)で設定してください", vbExclamation, "ERROR"
            Exit Sub
        End If
    End If
    If OptionText(1).Text = "" Then
        MsgBox "読出しｻｲｽﾞを入力してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If CLng(OptionText(1).Text) <= 0 Then
        MsgBox "読出しｻｲｽﾞが範囲外です", vbExclamation, "ERROR"
        Exit Sub
    End If
    
    'MACｱﾄﾞﾚｽ初期化有り
    If Check1.Value = vbChecked Then
        If OptionText(2).Text = "" Then
            MsgBox "MAC先頭番地を入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
    End If

    'ﾌｧｲﾙを保存するﾀﾞｲｱﾛｸﾞ
    SaveWorkFolder = CurDir             '現在のﾃﾞｨﾚｸﾄﾘ取得
    CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName = "" Then
        'MsgBox "ｷｬﾝｾﾙされました。"
        Exit Sub
    End If
    If tmpfso.FileExists(CommonDialog1.FileName) = True Then
        res = MsgBox("既にﾌｧｲﾙ" + CommonDialog1.FileTitle + "が存在します。上書きしますか？", vbQuestion + vbYesNo, "Overwrite?")
        If res = vbNo Then
            ChDir SaveWorkFolder        'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので元に戻す
            Exit Sub
        Else
            ReadMode = Overwrite_Read   '上書きﾓｰﾄﾞ
        End If
    End If
    ChDir SaveWorkFolder                'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので元に戻す
    
    'ROM読出しｵﾌﾟｼｮﾝ情報をｾｯﾄ
    ReadFileName = CommonDialog1.FileTitle
    ReadFilePass = CommonDialog1.FileName
    ReadAddress = CLng("&H" + OptionText(0).Text)
    If ReadAddress < 0 Then
        ReadAddress = ReadAddress + 65536
    End If
    ReadSize = CLng(OptionText(1).Text) * 1024
    If Check1.Value = vbChecked Then
        MacReadAddress = Val("&H" + OptionText(2).Text) + 1
        If MacReadAddress < 0 Then
            MacReadAddress = MacReadAddress + &H10000
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

    ReadMode = Cancel_Read
    Unload Me
    
End Sub



