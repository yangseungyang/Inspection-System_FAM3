VERSION 5.00
Begin VB.Form WriteSet_Form1 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Write"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   3480
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "："
      Height          =   255
      Left            =   1200
      TabIndex        =   18
      Top             =   585
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "："
      Height          =   255
      Left            =   1200
      TabIndex        =   17
      Top             =   585
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   1700
      TabIndex        =   16
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   1700
      TabIndex        =   15
      Top             =   525
      Width           =   135
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   1700
      TabIndex        =   11
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "："
      Height          =   255
      Left            =   1700
      TabIndex        =   7
      Top             =   1425
      Width           =   135
   End
   Begin VB.Label SumValue_lbl 
      Caption         =   "SumValue"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1425
      Width           =   1335
   End
   Begin VB.Label WriteFile_lbl 
      Caption         =   "WriteFile"
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   1125
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "："
      Height          =   255
      Left            =   1700
      TabIndex        =   4
      Top             =   1125
      Width           =   135
   End
   Begin VB.Label Label15 
      Caption         =   "ｻﾑ値"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1425
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "書込みﾌｧｲﾙ"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1125
      Width           =   975
   End
End
Attribute VB_Name = "WriteSet_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  WriteSet_Form1                                      *
'*                                                      *
'*      作成日  ：2007/02/22                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM書込み ｵﾌﾟｼｮﾝ画面（ﾉｰﾏﾙ）          *
'*                                                      *
'********************************************************


'--------------------------------------------------------
'
'   ROM書込みｵﾌﾟｼｮﾝ画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()
   
    ConnectAddress_lbl.Caption = Main_Form.ConnectTarget_cmb.Text    '接続先
    SlotNo_lbl.Caption = SelectedSlot                                'ｽﾛｯﾄ番号
    ModuleName_lbl.Caption = Main_Form.ModuleName(SelectedSlot - 1)  'ﾓｼﾞｭｰﾙ名
    WriteFile_lbl.Caption = Main_Form.FileName_txt.Text              '書込みﾌｧｲﾙ名
    SumValue_lbl.Caption = Main_Form.SumValue_lbl.Caption            '計算ｻﾑ値
    
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
'   OKﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub OK_btn_Click()
    
    'ROM書込みｵﾌﾟｼｮﾝ情報
    WriteAddress = 0                            '書込み開始ｱﾄﾞﾚｽ
    WriteSize = FileLen(SelectedFilePass)       '選択ﾌｧｲﾙのｻｲｽﾞを取得

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

