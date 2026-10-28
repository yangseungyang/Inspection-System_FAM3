VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ReadSet_Form1 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Read"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4155
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "読出しｵﾌﾟｼｮﾝ"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   2040
         MaxLength       =   7
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "ｴﾝﾃﾞｨｱﾝ"
         Height          =   615
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   2895
         Begin VB.OptionButton EndianSelect 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   17
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton EndianSelect 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   16
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label10 
            Caption         =   "Big"
            Height          =   255
            Left            =   840
            TabIndex        =   19
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "Little"
            Height          =   255
            Left            =   2040
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Label Label6 
         Caption         =   "読出しｱﾄﾞﾚｽ（HEX）"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   390
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "読出しｻｲｽﾞ（KByte）"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   750
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   240
      Width           =   135
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   525
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   825
      Width           =   135
   End
End
Attribute VB_Name = "ReadSet_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  ReadSet_Form                                        *
'*                                                      *
'*      作成日  ：2007/03/02                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM読出し ｵﾌﾟｼｮﾝ画面                  *
'*                                                      *
'********************************************************


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
    
    EndianSelect(0).Value = True
    Endian = BIG
    
    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    ReadAddress = 0
    ReadSize = 0
    
    'ROM読出し用ﾌﾗｸﾞ初期化
    ReadMode = Normal_Read
    InitMacAddress = DISABLE

End Sub


'--------------------------------------------------------
'
'   読出しｵﾌﾟｼｮﾝ
'
'--------------------------------------------------------
'--------------------------------------------------------
'   入力制限
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
    If CLng("&H" + OptionText(0).Text) Mod 512 <> 0 Then
        MsgBox "読出し開始ｱﾄﾞﾚｽは512Byte単位($200)で設定してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If OptionText(1).Text = "" Then
        MsgBox "読出しｻｲｽﾞを入力してください", vbExclamation, "ERROR"
        Exit Sub
    End If
    If Val(OptionText(1).Text) <= 0 Then
        MsgBox "読出しｻｲｽﾞが範囲外です", vbExclamation, "ERROR"
        Exit Sub
    End If

    'CMDGA特別ｴﾗｰﾁｪｯｸ
    If ModuleType = CMDGA_TYPE Then
        If Val("&H" + OptionText(0).Text) + (Val(OptionText(1).Text) * 1024) > 131072 Then
            MsgBox "CMDGAﾀｲﾌﾟのﾓｼﾞｭｰﾙはFlashROMの$20000番地以降へｱｸｾｽできません", vbExclamation, "ERROR"
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
    ReadAddress = Val("&H" + OptionText(0).Text)
    If ReadAddress < 0 Then
        ReadAddress = ReadAddress + &H10000
    End If
    ReadSize = Val(OptionText(1).Text) * 1024
    
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


