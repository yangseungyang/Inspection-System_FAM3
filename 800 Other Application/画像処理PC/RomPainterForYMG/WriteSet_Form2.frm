VERSION 5.00
Begin VB.Form WriteSet_Form2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Write"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5160
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   1320
      TabIndex        =   23
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2760
      TabIndex        =   22
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "書込みｵﾌﾟｼｮﾝ"
      Height          =   1215
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   4575
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   1
         Left            =   2640
         MaxLength       =   12
         TabIndex        =   21
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         IMEMode         =   3  'ｵﾌ固定
         Index           =   0
         Left            =   2640
         MaxLength       =   7
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   510
         Width           =   255
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   600
         Y1              =   450
         Y2              =   800
      End
      Begin VB.Line Line2 
         X1              =   600
         X2              =   720
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Line Line3 
         X1              =   600
         X2              =   720
         Y1              =   795
         Y2              =   795
      End
      Begin VB.Line Line4 
         X1              =   600
         X2              =   480
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label6 
         Caption         =   "MACｱﾄﾞﾚｽ"
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "MAC書込み番地(Hex)"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   400
         Width           =   1815
      End
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   525
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   825
      Width           =   1935
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "："
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label SumValue_lbl 
      Caption         =   "SumValue"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1425
      Width           =   1335
   End
   Begin VB.Label WriteFile_lbl 
      Caption         =   "WriteFile"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   1125
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "："
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1125
      Width           =   135
   End
   Begin VB.Label Label15 
      Caption         =   "ｻﾑ値"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1425
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "書込みﾌｧｲﾙ"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   1125
      Width           =   975
   End
End
Attribute VB_Name = "WriteSet_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*                                                          *
'*  WriteSet_Form2                                          *
'*                                                          *
'*      作成日  ：2007/03/15                                *
'*      作成者  ：澤原                                      *
'*      ｺﾒﾝﾄ    ：ROM書込み ｵﾌﾟｼｮﾝ画面（Ethernet,FL-net用） *
'*                                                          *
'************************************************************


'---------------------------------------------
'ｵﾌﾞｼﾞｪｸﾄ変数
'---------------------------------------------
Dim tmpFile As File
Dim tmpfso As New FileSystemObject


'--------------------------------------------------------
'
'   ROM書込みｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text       '接続先
    SlotNo_lbl.Caption = SelectedSlot                                   'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)     'ﾓｼﾞｭｰﾙ名
    WriteFile_lbl.Caption = Main_Form.FileName_txt.Text                 '書込みﾌｧｲﾙ名
    SumValue_lbl.Caption = Main_Form.SumValue_lbl.Caption               '計算ｻﾑ値

    OptionText(0).Enabled = True
    OptionText(0).BackColor = vbWhite
    OptionText(1).Enabled = True
    OptionText(1).BackColor = vbWhite
    
    Check1.Value = vbChecked
    
    OptionText(0).Text = "4000"
    
    'ｸﾞﾛｰﾊﾞﾙ変数初期化
    WriteAddress = 0
    WriteSize = 0
    Endian = BIG
    SaveFilePass = SelectedFilePass         '元のﾌｧｲﾙのﾊﾟｽを保存しておく

    'ROM書込み用ﾌﾗｸﾞ初期化
    WriteMode = Normal_Write
    WriteFirm = ENABLE

End Sub


'--------------------------------------------------------
'
'   書込みｵﾌﾟｼｮﾝ
'
'--------------------------------------------------------
'--------------------------------------------------------
'   MACｱﾄﾞﾚｽ書込みﾁｪｯｸﾎﾞｯｸｽ
'--------------------------------------------------------
Private Sub Check1_Click()

    If Check1.Value = vbChecked Then
        OptionText(0).Enabled = True
        OptionText(0).BackColor = vbWhite
        OptionText(1).Enabled = True
        OptionText(1).BackColor = vbWhite
    Else
        OptionText(0).Enabled = False
        OptionText(0).BackColor = &H8000000B
        OptionText(1).Enabled = False
        OptionText(1).BackColor = &H8000000B
    End If

End Sub

'--------------------------------------------------------
'   ﾃｷｽﾄ入力制限
'--------------------------------------------------------
Private Sub OptionText_KeyPress(Index As Integer, KeyAscii As Integer)

    '大文字入力のみ
    KeyCode = Chr(KeyAscii)
    KeyAscii = Asc(UCase(KeyCode))
    'BackSpaceとDelete有効
    If KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Sub
    'ｷｰ入力制限
    If KeyAscii < Asc("0") Or Asc("F") < KeyAscii Then
        KeyAscii = 0
    End If
    
End Sub


'--------------------------------------------------------
'
'   OKﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub OK_btn_Click()
    
    Dim i As Integer
    Dim tmpWriteFilePass As String
    Dim tmpFileNo As Integer

    '書込み開始ｱﾄﾞﾚｽ
    WriteAddress = 0
    
    '書込みｻｲｽﾞ=選択ﾌｧｲﾙのｻｲｽﾞを取得
    WriteSize = FileLen(SelectedFilePass)
    'MACｱﾄﾞﾚｽ書込み有りの場合
    If Check1.Value = vbChecked Then
        '---MAC書込み番地取得
        If OptionText(0).Text = "" Then
            MsgBox "MACｱﾄﾞﾚｽ書込み番地を入力してください", vbExclamation, "ERROR"
            Exit Sub
        End If
        If (CLng(OptionText(0).Text) < 0) Or (CLng("&H" + OptionText(0).Text) >= WriteSize - 12) Then
            MsgBox "MACｱﾄﾞﾚｽ書込み番地が範囲外です", vbExclamation, "ERROR"
            Exit Sub
        End If
        MacWriteAddress = CLng("&H" + OptionText(0).Text)
        If MacWriteAddress < 0 Then
            MacWriteAddress = MacWriteAddress + &H10000
        End If
        '---MACｱﾄﾞﾚｽ取得
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
        WriteMacData(6) = 0
        WriteMacData(7) = Not WriteMacData(3)
        WriteMacData(8) = Not WriteMacData(4)
        WriteMacData(9) = Not WriteMacData(5)
        '---MACｱﾄﾞﾚｽ付き書込みﾌｧｲﾙ作成
        tmpWriteFilePass = WorkFolder + "\tmpWriteFileWithMac.bin"  'MAC付書込み用ﾌｧｲﾙのﾌﾙﾊﾟｽ
        If tmpfso.FileExists(tmpWriteFilePass) Then                 'ﾌｧｲﾙがあったら削除
            Set tmpFile = tmpfso.GetFile(tmpWriteFilePass)
            tmpFile.Delete
        End If
        Set tmpFile = tmpfso.GetFile(SelectedFilePass)      '選択ﾌｧｲﾙに対応するFileｵﾌﾞｼﾞｪｸﾄを返す
        tmpFile.Copy (tmpWriteFilePass)                     '選択ﾌｧｲﾙをMAC付書込みﾌｧｲﾙにｺﾋﾟｰする
        Set tmpFile = Nothing
        Set tmpfso = Nothing
        tmpFileNo = FreeFile(0)                             '使用可能なﾌｧｲﾙ番号を取得
        Open tmpWriteFilePass For Binary As tmpFileNo       'MAC付書込みﾌｧｲﾙを開く
        Put tmpFileNo, MacWriteAddress + 1, WriteMacData()  'MACｱﾄﾞﾚｽをMAC付書込みﾌｧｲﾙに書込む
        Close tmpFileNo                                     'ﾌｧｲﾙを閉じる
        SelectedFilePass = tmpWriteFilePass
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
    Unload Me

End Sub

