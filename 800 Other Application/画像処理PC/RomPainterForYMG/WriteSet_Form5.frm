VERSION 5.00
Begin VB.Form WriteSet_Form5 
   Caption         =   "Form1"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton OK_btn 
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   26
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   495
      Left            =   2160
      TabIndex        =   25
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "書込みｵﾌﾟｼｮﾝ"
      Height          =   3855
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   3735
      Begin VB.CheckBox Check1 
         Height          =   420
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   285
         Width           =   255
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   0
         Left            =   1680
         TabIndex        =   17
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   420
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   420
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   765
         Width           =   255
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   1
         Left            =   1680
         TabIndex        =   13
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox OptionText 
         Height          =   270
         Index           =   4
         Left            =   2400
         TabIndex        =   11
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "システムプログラム"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "ブートプログラム"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "書込みｱﾄﾞﾚｽ"
         Height          =   255
         Left            =   960
         TabIndex        =   24
         Top             =   2685
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "ﾓｼﾞｭｰﾙ型名"
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   3030
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "MACｱﾄﾞﾚｽ"
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   3390
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "ｻﾑ値"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   2310
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "F/W書込み"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   1950
         Width           =   1215
      End
      Begin VB.Line Line4 
         X1              =   835
         X2              =   705
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line3 
         X1              =   720
         X2              =   720
         Y1              =   2280
         Y2              =   2760
      End
      Begin VB.Line Line1 
         X1              =   835
         X2              =   705
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label7 
         Caption         =   "0x"
         Height          =   255
         Left            =   2205
         TabIndex        =   19
         Top             =   2685
         Width           =   255
      End
   End
   Begin VB.Label Label5 
      Caption         =   "："
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   825
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "接続先"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "ｽﾛｯﾄ番号"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   525
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "ﾓｼﾞｭｰﾙ名"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "："
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   135
   End
   Begin VB.Label ConnectAddress_lbl 
      Caption         =   "ConnectAddress"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label SlotNo_lbl 
      Caption         =   "SlotNo"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   525
      Width           =   1335
   End
   Begin VB.Label ModuleName_lbl 
      Caption         =   "ModuleName"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "："
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   525
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "："
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   825
      Width           =   135
   End
End
Attribute VB_Name = "WriteSet_Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
