VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'なし
   ClientHeight    =   4725
   ClientLeft      =   6015
   ClientTop       =   4860
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Frame Frame1 
      Height          =   4890
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6120
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'なし
         Height          =   4815
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   4815
         ScaleWidth      =   6135
         TabIndex        =   1
         Top             =   0
         Width           =   6135
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   360
            Top             =   3720
         End
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    
    '二重起動の禁止
    If App.PrevInstance Then
        Timer1.Enabled = False
        MsgBox "二重起動はダメです！", vbCritical, "ERROR"
        Unload Me
        End
    End If

End Sub


Private Sub Timer1_Timer()
    
    Main_Form.Show
    Unload Me

End Sub
