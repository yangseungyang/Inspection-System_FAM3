VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main_Form 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Painter"
   ClientHeight    =   4335
   ClientLeft      =   2655
   ClientTop       =   2925
   ClientWidth     =   9525
   Icon            =   "Main_Form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   9525
   Begin VB.CommandButton Exit_btn 
      Caption         =   "終了"
      Height          =   495
      Left            =   8160
      TabIndex        =   52
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Verify_btn 
      Caption         =   "照合"
      Height          =   495
      Left            =   4200
      TabIndex        =   51
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Help_btn 
      Caption         =   "ヘルプ"
      Height          =   495
      Left            =   6840
      TabIndex        =   50
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Read_btn 
      Caption         =   "読出し"
      Height          =   495
      Left            =   5520
      TabIndex        =   49
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Copy_btn 
      Caption         =   "コピー"
      Height          =   495
      Left            =   8160
      TabIndex        =   48
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton EraseWrite_btn 
      Caption         =   "消去書込み"
      Height          =   495
      Left            =   4200
      TabIndex        =   47
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Erase_btn 
      Caption         =   "消去"
      Height          =   495
      Left            =   5520
      TabIndex        =   46
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Write_btn 
      Caption         =   "書込み"
      Height          =   495
      Left            =   6840
      TabIndex        =   45
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "ﾌｧｲﾙ選択"
      Height          =   1240
      Left            =   240
      TabIndex        =   39
      Top             =   2880
      Width           =   3720
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3000
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton FileOpen_btn 
         Caption         =   "参照"
         Height          =   350
         Left            =   3000
         TabIndex        =   41
         Top             =   270
         Width           =   520
      End
      Begin VB.TextBox FileName_txt 
         Height          =   300
         Left            =   960
         Locked          =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   40
         Top             =   300
         Width           =   1935
      End
      Begin VB.Label SumValue_lbl 
         BorderStyle     =   1  '実線
         Height          =   300
         Left            =   960
         TabIndex        =   44
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '透明
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '透明
         Caption         =   "ﾌｧｲﾙ名"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox SysLog_txt 
      Height          =   2160
      Left            =   6240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直
      TabIndex        =   37
      Top             =   435
      Width           =   3015
   End
   Begin VB.CommandButton LogClear_btn 
      Caption         =   "Clear"
      Height          =   255
      Left            =   8640
      TabIndex        =   36
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "接続先"
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.OptionButton CpuSelect_btn 
         Caption         =   "Option1"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   54
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton CpuSelect_btn 
         Caption         =   "Option1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   53
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton IOinfoUpdate_btn 
         Caption         =   "I/O構成更新"
         Height          =   350
         Left            =   2280
         TabIndex        =   19
         Top             =   330
         Width           =   1455
      End
      Begin VB.CommandButton SysReset_btn 
         Caption         =   "ｼｽﾃﾑﾘｾｯﾄ"
         Height          =   350
         Left            =   3960
         TabIndex        =   18
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   15
         Left            =   5145
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   17
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   14
         Left            =   4890
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   16
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   13
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   15
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   12
         Left            =   4380
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   14
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   11
         Left            =   4125
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   13
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   3870
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   12
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   3615
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   11
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   10
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   3105
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   9
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   2850
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   8
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   2595
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   7
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   6
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   5
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   1830
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   4
         Top             =   1200
         Width           =   250
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   4
         MousePointer    =   1  '矢印
         MultiLine       =   -1  'True
         OLEDropMode     =   1  '手動
         TabIndex        =   3
         Top             =   1200
         Width           =   250
      End
      Begin VB.ComboBox ConnectTarget_cmb 
         Height          =   300
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "100.100.100.100"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox ModuleName 
         Alignment       =   2  '中央揃え
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   4
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   250
      End
      Begin VB.Label Label2 
         Caption         =   "CPU2"
         Height          =   255
         Left            =   600
         TabIndex        =   56
         Top             =   1830
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "CPU1"
         Height          =   255
         Left            =   600
         TabIndex        =   55
         Top             =   1470
         Width           =   495
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         BorderWidth     =   5
         X1              =   720
         X2              =   1560
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   5
         X1              =   720
         X2              =   720
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "10"
         Height          =   255
         Index           =   9
         Left            =   3660
         TabIndex        =   35
         Top             =   960
         Width           =   255
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "9"
         Height          =   255
         Index           =   8
         Left            =   3450
         TabIndex        =   34
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "8"
         Height          =   255
         Index           =   7
         Left            =   3210
         TabIndex        =   33
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "7"
         Height          =   255
         Index           =   6
         Left            =   2955
         TabIndex        =   32
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "6"
         Height          =   255
         Index           =   5
         Left            =   2685
         TabIndex        =   31
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "5"
         Height          =   255
         Index           =   4
         Left            =   2430
         TabIndex        =   30
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "4"
         Height          =   255
         Index           =   3
         Left            =   2175
         TabIndex        =   29
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "3"
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   28
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "2"
         Height          =   255
         Index           =   1
         Left            =   1665
         TabIndex        =   27
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   1410
         TabIndex        =   26
         Top             =   960
         Width           =   135
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "11"
         Height          =   255
         Index           =   10
         Left            =   3915
         TabIndex        =   25
         Top             =   960
         Width           =   255
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "12"
         Height          =   255
         Index           =   11
         Left            =   4170
         TabIndex        =   24
         Top             =   960
         Width           =   255
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "13"
         Height          =   255
         Index           =   12
         Left            =   4425
         TabIndex        =   23
         Top             =   960
         Width           =   255
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "14"
         Height          =   255
         Index           =   13
         Left            =   4680
         TabIndex        =   22
         Top             =   960
         Width           =   255
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "15"
         Height          =   255
         Index           =   14
         Left            =   4920
         TabIndex        =   21
         Top             =   960
         Width           =   255
      End
      Begin VB.Label SlotNo 
         BackStyle       =   0  '透明
         Caption         =   "16"
         Height          =   255
         Index           =   15
         Left            =   5175
         TabIndex        =   20
         Top             =   960
         Width           =   375
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "ｼｽﾃﾑﾛｸﾞ"
      Height          =   255
      Left            =   6240
      TabIndex        =   38
      Top             =   180
      Width           =   735
   End
   Begin VB.Menu MyPopMe 
      Caption         =   "右ｸﾘｯｸ時のﾎﾟｯﾌﾟｱｯﾌﾟﾒﾆｭｰ"
      Visible         =   0   'False
      Begin VB.Menu Sys1 
         Caption         =   "SYS1(PF,KPｹﾞｰﾄｱﾚｲ使用ﾓｼﾞｭｰﾙ)"
      End
      Begin VB.Menu Sys2 
         Caption         =   "SYS2(ｺﾏﾝﾄﾞｹﾞｰﾄｱﾚｲ使用ﾓｼﾞｭｰﾙ)"
      End
      Begin VB.Menu Sys3 
         Caption         =   "SYS3(DeviceNetﾓｼﾞｭｰﾙ)"
      End
      Begin VB.Menu Sys4 
         Caption         =   "SYS4(FL-netﾓｼﾞｭｰﾙ)"
      End
      Begin VB.Menu Sys5 
         Caption         =   "SYS5(RTOS-CPUﾓｼﾞｭｰﾙ)"
      End
      Begin VB.Menu Sys6 
         Caption         =   "SYS6(新ｼﾘｱﾙﾓｼﾞｭｰﾙ)"
      End
      Begin VB.Menu Sys7 
         Caption         =   "SYS7(Ethernetﾓｼﾞｭｰﾙ)"
      End
      Begin VB.Menu Sys8 
         Caption         =   "SYS8(新Ethernetﾓｼﾞｭｰﾙ)"
      End
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  Main_Form                                           *
'*                                                      *
'*      作成日  ：2007/02/14                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM Painter ﾒｲﾝ画面                   *
'*                                                      *
'********************************************************

'Option Explicit

'---------------------------------------------
'外部API関数使用宣言
'---------------------------------------------
'---ｷｬﾚｯﾄを消す
Private Declare Function HideCaret Lib "user32.dll" (ByVal hwnd As Long) As Long
'---ｷｬﾚｯﾄを表示する
Private Declare Function ShowCaret Lib "user32.dll" (ByVal hwnd As Long) As Long
'---Sleep関数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------
'ｵﾌﾞｼﾞｪｸﾄ変数の宣言
'---------------------------------------------
'---FileSystemObjectｵﾌﾞｼﾞｪｸﾄ
Private fso As New FileSystemObject



'--------------------------------------------------------
'
'   初期化処理
'
'--------------------------------------------------------
Private Sub Form_Load()
    
    Dim dummy As String
    Dim InitTargetIndex, InitSlotNo As Integer
    
    '機能ﾛｯｸ設定ﾌｧｲﾙ
    '---ﾌｧｲﾙがなかったら作成
    If fso.FileExists("FunctionLock.txt") = False Then
        Set FunctionLock = fso.CreateTextFile("FunctionLock.txt", True)
        FunctionLock.WriteLine ("---機能ﾛｯｸ設定ﾌｧｲﾙ---")
        FunctionLock.WriteLine ("EraseWrite = TRUE")
        FunctionLock.WriteLine ("Erase = TRUE")
        FunctionLock.WriteLine ("Write = TRUE")
        FunctionLock.WriteLine ("Copy = TRUE")
        FunctionLock.WriteLine ("Check = TRUE")
        FunctionLock.WriteLine ("Read = TRUE")
        FunctionLock.WriteLine ("Help = TRUE")
        FunctionLock.Close
    End If
    '---ﾌｧｲﾙを開く
    Set FunctionLock = fso.OpenTextFile("FunctionLock.txt", ForReading)
    dummy = FunctionLock.ReadLine               '一行目は無視
    Do While FunctionLock.AtEndOfLine = False   'ﾌｧｲﾙﾎﾟｲﾝﾀがEOLの直前にいくまでﾙｰﾌﾟ
        ReadLineBuff = FunctionLock.ReadLine    '一行読出し ﾎﾟｲﾝﾀ+1
        Select Case ReadLineBuff
        Case "EraseWrite = FALSE"
            EraseWrite_btn.Enabled = False      '消去書込み機能ﾛｯｸ
        Case "Erase = FALSE"
            Erase_btn.Enabled = False           '消去機能ﾛｯｸ
        Case "Write = FALSE"
            Write_btn.Enabled = False           '書込み機能ﾛｯｸ
        Case "Copy = FALSE"
            Copy_btn.Enabled = False            'ｺﾋﾟｰ機能ﾛｯｸ
        Case "Check = FALSE"
            Verify_btn.Enabled = False           '照合機能ﾛｯｸ
        Case "Read = FALSE"
            Read_btn.Enabled = False            '読出し機能ﾛｯｸ
        Case "Help = FALSE"
            Help_btn.Enabled = False            'ﾍﾙﾌﾟ機能ﾛｯｸ
        End Select
    Loop
    '---ﾌｧｲﾙを閉じる
    FunctionLock.Close
    

    '設定値保存ﾌｧｲﾙ処理
    '---ﾌｧｲﾙがなかったら作成
    If fso.FileExists("SaveValues.txt") = False Then
        Set SaveValues = fso.CreateTextFile("SaveValues.txt", True)
        SaveValues.WriteLine ("---設定値保存ﾌｧｲﾙ---")
        SaveValues.WriteLine (Str(0))
        SaveValues.WriteLine (Str(3))
        SaveValues.WriteLine (CurDir)
        SaveValues.Close
    End If
    '---ﾌｧｲﾙを開く
    Set SaveValues = fso.OpenTextFile("SaveValues.txt", ForReading)
    '---読出し
    dummy = SaveValues.ReadLine                      '1行目は無視
    InitTargetIndex = Val(SaveValues.ReadLine)       '2行目(接続先初期値)
    InitSlotNo = Val(SaveValues.ReadLine)            '3行目(ｽﾛｯﾄ位置初期値)
    WorkFolder = SaveValues.ReadLine                 '4行目(ｲﾝｽﾄｰﾙされているﾊﾟｽ)
    '---ﾌｧｲﾙを閉じる
    SaveValues.Close

    
    '接続先ｱﾄﾞﾚｽ保存ﾌｧｲﾙ
    '---ﾌｧｲﾙがなかったら作成
    If fso.FileExists("ConnectAddress.txt") = False Then
        Set ConnectAddress = fso.CreateTextFile("ConnectAddress.txt", True)
        ConnectAddress.WriteLine ("---接続先ｱﾄﾞﾚｽ保存ﾌｧｲﾙ---")
        ConnectAddress.WriteLine ("COM1")
        ConnectAddress.WriteLine ("COM2")
        ConnectAddress.Close
    End If
    '---ﾌｧｲﾙを開く
    Set ConnectAddress = fso.OpenTextFile("ConnectAddress.txt", ForReading)
    ConnectTarget_cmb.Clear                                     '接続先項目ｸﾘｱ
    '---読出し
    dummy = ConnectAddress.ReadLine                             '一行目は無視
    Do While ConnectAddress.AtEndOfLine = False
        ConnectTarget_cmb.AddItem (ConnectAddress.ReadLine)     '接続先に項目追加
    Loop
    '---ﾌｧｲﾙを閉じる
    ConnectAddress.Close
    
    'Call ChDir("C:\local\work\FA-M3 Tools\RomPainter")
    'Debug.Print WorkFolder
    
    '変数初期化
    SelectedSlot = CInt(InitSlotNo)
    DefaultMAC(0) = &H0                 'ﾃﾞﾌｫﾙﾄMACｱﾄﾞﾚｽ(ｴﾝﾃﾞｨｱﾝ注意)
    DefaultMAC(1) = &HFF64
    DefaultMAC(2) = &HFFFF
    DefaultMAC(3) = &H0
    DefaultMAC(4) = &H0

    
    'ﾌｫｰﾑ画面初期設定
    '---選択CPUの初期値ｾｯﾄ
    CpuSelect_btn(1).Value = True
    '---接続先の初期値ｾｯﾄ
    ConnectTarget_cmb.ListIndex = InitTargetIndex
    '---指定ｽﾛｯﾄの初期値ｾｯﾄ
    ModuleName(SelectedSlot - 1).BackColor = RGB(255, 179, 181)


    'ｸﾗｽﾓｼﾞｭｰﾙ:ComMng_Modeleを使用する宣言
    Set cmm = New ComMng_Module
    '---ComMng_Modeleﾌﾟﾛﾊﾟﾃｨ設定
    cmm.AppName = "RomPainter"
    '---通信ﾗｲﾌﾞﾗﾘ使用開始
    Call cmm.ComCtrlLibInit
    
End Sub



'--------------------------------------------------------
'
'   終了処理
'
'--------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)

    Dim TargetIndex As Integer

    '接続先ｱﾄﾞﾚｽ保存ﾌｧｲﾙに上書き
    Set ConnectAddress = fso.OpenTextFile(App.Path & "\ConnectAddress.txt", ForWriting)
    ConnectAddress.WriteLine ("---接続先ｱﾄﾞﾚｽ保存ﾌｧｲﾙ---")
    For i = 0 To ConnectTarget_cmb.ListCount - 1
        If ConnectTarget_cmb.List(i) = ConnectTarget_cmb.Text Then
            TargetIndex = i    '設定保存ﾌｧｲﾙ用
        End If
        ConnectAddress.WriteLine (ConnectTarget_cmb.List(i))
    Next
    ConnectAddress.Close
    
    '設定値保存ﾌｧｲﾙに上書き
    Set SaveValues = fso.OpenTextFile(App.Path & "\SaveValues.txt", ForWriting)
    SaveValues.WriteLine ("---設定保存ﾌｧｲﾙ---")
    SaveValues.WriteLine (Str(TargetIndex))
    SaveValues.WriteLine (Str(SelectedSlot))
    SaveValues.WriteLine (WorkFolder)
    SaveValues.Close
    
    
    'ﾌｧｲﾙｼｽﾃﾑｵﾌﾞｼﾞｪｸﾄｸﾗｽの解放
    Set SaveValues = Nothing
    Set fso = Nothing
    
    '通信ﾗｲﾌﾞﾗﾘ(ComCtrlEx.dll)使用終了
    cmm.ComCtrlLibEnd
        
    'ｸﾗｽﾓｼﾞｭｰﾙ:ComMng_Modeleの解放
    Set cmm = Nothing

End Sub



'--------------------------------------------------------
'
'   ﾌｫｰﾑ関連の処理
'
'--------------------------------------------------------
'---ｳｨﾝﾄﾞｳが切り替わった時の処理
Private Sub Form_Paint()

    Dim i As Integer

    For i = 0 To 15
        HideCaret (ModuleName(i).hwnd)  'ﾓｼﾞｭｰﾙ名ﾃｷｽﾄのｷｬﾚｯﾄ非表示
    Next

    HideCaret (SysLog_txt.hwnd)         'ｼｽﾃﾑﾛｸﾞのｷｬﾚｯﾄ非表示

End Sub

'---ﾓｼﾞｭｰﾙ名にﾌｫｰｶｽが設定された時の処理
Private Sub ModuleName_GotFocus(Index As Integer)

    HideCaret (ModuleName(Index).hwnd)  'ｷｬﾚｯﾄ非表示
    ModuleName(Index).Locked = True     'ﾃｷｽﾄ入力禁止

End Sub

'---ﾓｼﾞｭｰﾙ名がﾀﾞﾌﾞﾙｸﾘｯｸされた時の処理
Private Sub ModuleName_DblClick(Index As Integer)

    'ﾃｷｽﾄ入力許可ﾓｰﾄﾞへ
    If Index <> 0 Then                      'ｽﾛｯﾄ1は無視
        ShowCaret (ModuleName(Index).hwnd)  'ｷｬﾚｯﾄ表示
        ModuleName(Index).Locked = False    'ﾃｷｽﾄ入力許可
        ModuleName(Index).IMEMode = 3       '半角英数字固定
        ShowCaret (ModuleName(Index).hwnd)  'ｷｬﾚｯﾄ表示
    End If

End Sub

'---ﾓｼﾞｭｰﾙ名で何かｷｰを押された時の処理
Private Sub ModuleName_KeyPress(Index As Integer, KeyAscii As Integer)

    Dim KeyCode As String

    'ｽﾍﾟｰｽｷｰを無視する
    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
    End If

    'ENTERｷｰが押されるとﾃｷｽﾄ入力禁止ﾓｰﾄﾞへ
    If KeyAscii = vbKeyReturn Then
        HideCaret (ModuleName(Index).hwnd)  'ｷｬﾚｯﾄ非表示
        ModuleName(Index).Locked = True     'ﾃｷｽﾄ入力禁止
    End If

    '大文字入力のみ
    KeyCode = Chr(KeyAscii)
    KeyAscii = Asc(UCase(KeyCode))

End Sub

'---ﾓｼﾞｭｰﾙ名がｸﾘｯｸされた時の処理
Private Sub ModuleName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    If CpuSelect_btn(0).Value = True Then
        If Index <> 0 Then                                      'ｽﾛｯﾄ1は無視
            For i = 1 To 15                                     '一旦
                ModuleName(i).BackColor = vbWhite               ' 全部白
            Next                                                '　にして、
            ModuleName(Index).BackColor = RGB(255, 179, 181)    'ｸﾘｯｸされたとこだけ赤
            SelectedSlot = Index + 1                            '選択ｽﾛｯﾄ番号
        End If
    Else
        If Index > 1 Then                                      'ｽﾛｯﾄ1は無視
            For i = 2 To 15                                     '一旦
                ModuleName(i).BackColor = vbWhite               ' 全部白
            Next                                                '　にして、
            ModuleName(Index).BackColor = RGB(255, 179, 181)    'ｸﾘｯｸされたとこだけ赤
            SelectedSlot = Index + 1                            '選択ｽﾛｯﾄ番号
        End If
    End If
    
    ModuleName(Index).Enabled = False
    ModuleName(Index).Enabled = True
    ModuleName(Index).SetFocus
    
    'ﾏｳｽの右ｸﾘｯｸを検地
    If Button = vbRightButton Then
        'ﾎﾟｯﾌﾟｱｯﾌﾟﾒﾆｭｰの表示
        Main_Form.PopupMenu MyPopMe
    End If

End Sub

'---ｼｽﾃﾑﾛｸﾞにﾌｫｰｶｽが設定された時の処理
Private Sub SysLog_txt_GotFocus()

    HideCaret (SysLog_txt.hwnd)

End Sub

'---ﾌｧｲﾙ選択にﾌｫｰｶｽが設定された時の処理
Private Sub FileName_txt_GotFocus()

    HideCaret (FileName_txt.hwnd)

End Sub

'---CPU Number選択
Private Sub CpuSelect_btn_Click(Index As Integer)

    If Index = 0 Then
        CpuNo = CPU1
        ModuleName(1).BackColor = vbWhite
        Line2.X2 = 1320
    Else
        CpuNo = CPU2
        ModuleName(1).BackColor = &H80000003
        Line2.X2 = 1560
        If SelectedSlot = 2 Then
            SelectedSlot = 3
            ModuleName(SelectedSlot - 1).BackColor = RGB(255, 179, 181)
        End If
    End If

End Sub



'--------------------------------------------------------
'
'   ﾛｸﾞｸﾘｱﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub LogClear_btn_Click()

    SysLog_txt.Text = ""

End Sub



'--------------------------------------------------------
'
'   参照ﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub FileOpen_btn_Click()

    Dim SaveWorkFolder As String

    SaveWorkFolder = CurDir      '現在のﾃﾞｨﾚｸﾄﾘ取得

    CommonDialog1.Filter = "binファイル(*.bin)|*.bin|" & "全てのファイル(*.*)|*.*"
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName = "" Then
        'MsgBox ("キャンセルされました。")
        Exit Sub
    Else
        FileName_txt.Text = CommonDialog1.FileTitle     '選択ﾌｧｲﾙ名
        SelectedFilePass = CommonDialog1.FileName       '選択ﾌｧｲﾙのﾌﾙﾊﾟｽ
    End If
    
    ChDir SaveWorkFolder     'ｺﾓﾝﾀﾞｲｱﾛｸﾞでﾃﾞｨﾚｸﾄﾘが変更されるので、元に戻す
    
    'binﾌｧｲﾙのｻﾑ値表示
    FileSize = FileLen(SelectedFilePass)            '選択ﾌｧｲﾙのｻｲｽﾞを取得
    FileNo = FreeFile(0)                            '使用可能なﾌｧｲﾙ番号を取得
    Open SelectedFilePass For Binary As FileNo      'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
    Get FileNo, FileSize - 1, SumValue()            '選択ﾌｧｲﾙのｻﾑ値を取得
    CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
    If CalSumValue > 65535 Then
        CalSumValue = CalSumValue - 65535
    End If
    SumValue_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    Close FileNo                                    'ﾊﾞｲﾅﾘﾌｧｲﾙを閉じる

End Sub



'--------------------------------------------------------
'
'   I/O構成更新ﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub IOinfoUpdate_btn_Click()

    FAM3Connect                                     'FAM3と接続
    If cmm.ConnectState = Connecting Then           '接続できたら、
        Call cmm.GetIOinfo(IOinfo())                '　→I/O構成情報取得
        If cmm.CmdState = CMD_OK Then               '　→取得できたら、
            For i = 0 To 15                         '　　→I/O構成更新
                ModuleName(i) = IOinfo(i)
            Next i
            AddSysLog ("I/O構成更新")
        Else
            AddSysLog ("I/O構成の取得に失敗しました")
        End If
        FAM3DisConnect                              'FAM3から切断
    End If

End Sub



'--------------------------------------------------------
'
'   ｼｽﾃﾑﾘｾｯﾄﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub SysReset_btn_Click()

    If CpuNo = CPU2 Then
        MsgBox "アドオンCPUではリセットスタートできません", vbCritical, "EOOR"
        Exit Sub
    End If
        
    cmm.ConnectState = Reset
    
    FAM3Connect                             'FAM3と接続
    
    If cmm.ConnectState = Connecting Then   '接続できたら、
        cmm.ConnectState = Reset
        Call cmm.SystemReset(-1)            '　→全CPUﾘｽﾀｰﾄ
        FAM3DisConnect                      '　→FAM3から切断
        Sleep (1000)
        MsgBox "リセットスタート完了しました", vbInformation, "SUCCESS"
        AddSysLog ("ﾘｾｯﾄｽﾀｰﾄ完了")
    End If
    
    m_ConnectState = NotConnect             'ｽﾃｰﾀｽを強制的に未接続にする

End Sub



'--------------------------------------------------------
'
'   ROM消去書込みﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub EraseWrite_btn_Click()

    If ErrorCheck(1, 1, 1, 1) = NG Then Exit Sub    'ｴﾗｰﾁｪｯｸ
    
    Call Erase_btn_Click
    
    If EraseMode <> Cancel_Erase Then               'ｷｬﾝｾﾙ以外なら、
        Call Write_btn_Click                        'ROM書込みへ
    End If

End Sub



'--------------------------------------------------------
'
'   ROM消去ﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Erase_btn_Click()

    If ErrorCheck(1, 1, 1, 0) = NG Then Exit Sub    'ｴﾗｰﾁｪｯｸ
    
    Select Case ModuleType                          'ROM消去ｵﾌﾟｼｮﾝ画面
        Case PFKP_TYPE, CMDGA_TYPE, DEVNET_TYPE, FLNET_TYPE, RTOS_TYPE, NSERIAL_TYPE, ETHER_TYPE, NETHER_TYPE, AROUSAL_TYPE
            EraseSet_Form.Show vbModal
        Case Else
            MsgBox "このモジュールのROM消去はサポートしていません", vbCritical, "SORRY"
            EraseMode = Cancel_Erase
    End Select
    
    If EraseMode <> Cancel_Erase Then               'ｷｬﾝｾﾙ以外なら、
        Call FAM3Connect                            '　→FAM3と接続
        If (cmm.ConnectState = Connecting) Then     '接続できたら、
            EraseRun_Form.Show vbModal              '　→ROM消去画面へ
            Call FAM3DisConnect                     '　→FAM3から切断
        End If
    End If
    
End Sub



'--------------------------------------------------------
'
'   ROM書込みﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Write_btn_Click()

    If ErrorCheck(1, 1, 1, 1) = NG Then Exit Sub    'ｴﾗｰﾁｪｯｸ
    
    Select Case ModuleType                          'ROM書込ｵﾌﾟｼｮﾝ画面へ
        Case CMDGA_TYPE, NSERIAL_TYPE, PFKP_TYPE
            WriteSet_Form1.Show vbModal
        Case FLNET_TYPE, ETHER_TYPE, NETHER_TYPE
            WriteSet_Form2.Show vbModal
        Case RTOS_TYPE
            WriteSet_Form3.Show vbModal
        Case DEVNET_TYPE
            WriteSet_Form4.Show vbModal
        Case Else
            MsgBox "このモジュールのROM書込みはサポートしていません", vbCritical, "SORRY"
            WriteMode = Cancel_Write
    End Select
    
    If WriteMode <> Cancel_Write Then               'ｷｬﾝｾﾙ以外なら、
        FAM3Connect                                 '　→FAM3と接続
        If cmm.ConnectState = Connecting Then       '接続できたら、
            WriteRun_Form.Show vbModal              '　→ROM書込み画面へ
            FAM3DisConnect                          '　→FAM3から切断
        End If
    End If

End Sub



'--------------------------------------------------------
'
'   ROMｺﾋﾟｰﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Copy_btn_Click()

    Dim m_SP As String

    m_SP = Left$(Main_Form.ModuleName(0).Text, 4)
    If (m_SP = "SP21") Or (m_SP = "SP25") Or (m_SP = "SP35") Then
        MsgBox "ROMコピーはこのCPUではできません", vbExclamation, "SORRY!"
        Exit Sub
    End If
        
    If ErrorCheck(1, 1, 1, 0) = NG Then Exit Sub    'ｴﾗｰﾁｪｯｸ
    
    Select Case ModuleType                          'ROMｺﾋﾟｰｵﾌﾟｼｮﾝ画面へ
        Case PFKP_TYPE, CMDGA_TYPE, DEVNET_TYPE, NSERIAL_TYPE, RTOS_TYPE
            CopySet_Form1.Show vbModal
        Case ETHER_TYPE, NETHER_TYPE
            CopySet_Form2.Show vbModal
        Case Else
            MsgBox "このモジュールのROMコピーはサポートしていません", vbCritical, "SORRY"
            CopyMode = Cancel_Copy
    End Select
    
    If CopyMode <> Cancel_Copy Then                 'ｷｬﾝｾﾙ以外なら、
        FAM3Connect                                 '　→FAM3と接続
        If cmm.ConnectState = Connecting Then       '接続できたら、
            CopyRun_Form.Show vbModal               '　→ROMｺﾋﾟｰ画面へ
            FAM3DisConnect                          '　→FAM3から切断
        End If
    End If

End Sub



'--------------------------------------------------------
'
'   ROM照合ﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Verify_btn_Click()

    If ErrorCheck(1, 1, 1, 1) = NG Then Exit Sub    'ｴﾗｰﾁｪｯｸ
    
    Select Case ModuleType                          'ROM照合ｵﾌﾟｼｮﾝ画面へ
        Case PFKP_TYPE, CMDGA_TYPE, NSERIAL_TYPE
            VerifySet_Form1.Show vbModal
        Case FLNET_TYPE, ETHER_TYPE, NETHER_TYPE
            VerifySet_Form2.Show vbModal
        Case RTOS_TYPE
            VerifySet_Form3.Show vbModal
        Case DEVNET_TYPE
            VerifySet_Form4.Show vbModal
        Case Else
            MsgBox "このモジュールのROM照合はサポートしていません", vbCritical, "SORRY"
            VerifyMode = Cancel_Verify
    End Select
    
    If VerifyMode <> Cancel_Verify Then             'ｷｬﾝｾﾙ以外なら、
        FAM3Connect                                 '　→FAM3と接続
        If cmm.ConnectState = Connecting Then       '接続できたら、
            VerifyRun_Form.Show vbModal             '　→ROM照合画面へ
            FAM3DisConnect                          '　→FAM3から切断
        End If
    End If

End Sub



'--------------------------------------------------------
'
'   ROM読出しﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Read_btn_Click()

    If ErrorCheck(1, 1, 1, 0) = NG Then Exit Sub    'ｴﾗｰﾁｪｯｸ

    Select Case ModuleType                          'ROM読出しｵﾌﾟｼｮﾝ画面へ
        Case PFKP_TYPE, CMDGA_TYPE, DEVNET_TYPE, NSERIAL_TYPE, RTOS_TYPE
            ReadSet_Form1.Show vbModal
        Case FLNET_TYPE, ETHER_TYPE, NETHER_TYPE
            ReadSet_Form2.Show vbModal
        Case Else
            MsgBox "このモジュールのROM読出しはサポートしていません", vbCritical, "SORRY"
            ReadMode = Cancel_Read
    End Select
    
    If ReadMode <> Cancel_Read Then                 'ｷｬﾝｾﾙ以外なら、
        Call FAM3Connect                            '　→FAM3と接続
        If (cmm.ConnectState = Connecting) Then     '接続できたら、
            ReadRun_Form.Show vbModal               '　→ROM読出し画面へ
            Call FAM3DisConnect                     '　→FAM3から切断
        End If
    End If

End Sub



'--------------------------------------------------------
'
'   終了ﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Exit_btn_Click()

    Unload Main_Form

End Sub



'--------------------------------------------------------
'
'   ﾍﾙﾌﾟﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Help_btn_Click()
    
    SendKeys "{F1}"

End Sub



'--------------------------------------------------------
'
'   拡張機能
'
'--------------------------------------------------------
'--------------------------------------------------------
'   選択ﾌｧｲﾙﾃｷｽﾄﾎﾞｯｸｽにﾌｧｲﾙがﾄﾞﾛｯﾌﾟされた場合
'--------------------------------------------------------
Private Sub FileName_txt_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim yenpos As Integer

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
    SelectedFilePass = Data.Files(1)
            
    '\を後ろから検索して文字のある場所を文字数で返す
    yenpos = InStrRev(SelectedFilePass, "\")
    'ﾌｧｲﾙ名を取得して表示
    FileName_txt.Text = Right$(SelectedFilePass, Len(SelectedFilePass) - yenpos)
            
    'binﾌｧｲﾙのｻﾑ値表示
    FileSize = FileLen(SelectedFilePass)            '選択ﾌｧｲﾙのｻｲｽﾞを取得
    FileNo = FreeFile(0)                            '使用可能なﾌｧｲﾙ番号を取得
    Open SelectedFilePass For Binary As FileNo      'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
    Get FileNo, FileSize - 1, SumValue()            '選択ﾌｧｲﾙのｻﾑ値を取得
    CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
    If CalSumValue > 65535 Then
        CalSumValue = CalSumValue - 65535
    End If
    SumValue_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    Close FileNo

End Sub

'--------------------------------------------------------
'   ﾓｼﾞｭｰﾙ名ﾃｷｽﾄﾎﾞｯｸｽにﾌｧｲﾙがﾄﾞﾛｯﾌﾟされたらROM書込み処理
'--------------------------------------------------------
Private Sub ModuleName_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    Dim yenpos As Integer

    '選択ｽﾛｯﾄ
    If Index <> 0 Then                                      'ｽﾛｯﾄ1は無視
        For i = 1 To 15                                     '一旦
            ModuleName(i).BackColor = vbWhite               ' 全部白
        Next                                                '　にして、
        ModuleName(Index).BackColor = RGB(255, 179, 181)    'ｸﾘｯｸされたとこだけ赤
        SelectedSlot = Index + 1                            '選択ｽﾛｯﾄ番号
    End If

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
    SelectedFilePass = Data.Files(1)
            
    '\を後ろから検索して文字のある場所を文字数で返す
    yenpos = InStrRev(SelectedFilePass, "\")
    'ﾌｧｲﾙ名を取得して表示
    FileName_txt.Text = Right$(SelectedFilePass, Len(SelectedFilePass) - yenpos)
            
    'binﾌｧｲﾙのｻﾑ値表示
    FileSize = FileLen(SelectedFilePass)            '選択ﾌｧｲﾙのｻｲｽﾞを取得
    FileNo = FreeFile(0)                            '使用可能なﾌｧｲﾙ番号を取得
    Open SelectedFilePass For Binary As FileNo      'ﾊﾞｲﾅﾘﾌｧｲﾙを開く
    Get FileNo, FileSize - 1, SumValue()            '選択ﾌｧｲﾙのｻﾑ値を取得
    CalSumValue = (Val(SumValue(0)) * &H100) + SumValue(1) + SumValue(0) + SumValue(1)
    If CalSumValue > 65535 Then
        CalSumValue = CalSumValue - 65535
    End If
    SumValue_lbl.Caption = Right$("0000" & Hex(CalSumValue), 4)
    Close FileNo

    '書込みﾎﾞﾀﾝが押された時の処理へ
    Call Write_btn_Click
    
End Sub

'--------------------------------------------------------
'   右ｸﾘｯｸﾒﾆｭｰがｸﾘｯｸされた時の処理
'--------------------------------------------------------
Private Sub Sys1_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS1"
End Sub
Private Sub Sys2_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS2"
End Sub
Private Sub Sys3_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS3"
End Sub
Private Sub Sys4_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS4"
End Sub
Private Sub Sys5_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS5"
End Sub
Private Sub Sys6_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS6"
End Sub
Private Sub Sys7_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS7"
End Sub
Private Sub Sys8_Click()
    ModuleName(SelectedSlot - 1).Text = "SYS8"
End Sub

