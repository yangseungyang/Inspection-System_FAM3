VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CopyRun_Form 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Copy Running"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4005
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "中断"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   41
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Rate 
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "CopyRun_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  CopyRun_Form                                        *
'*                                                      *
'*      作成日  ：2007/06/21                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROMｺﾋﾟｰ 実行画面                      *
'*                                                      *
'********************************************************

'Option Explicit

'---------------------------------------------
'外部API関数使用宣言
'---------------------------------------------
'---Sleep関数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'--------------------------------------------------------
'
'   ROMｺﾋﾟｰ実行画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    'ROMｺﾋﾟｰ動画ｽﾀｰﾄ
    Animation1.Open WorkFolder + "\CopyAnimation.avi"
    Animation1.Play -1
    
    AllCopySize = CopySize                      'ｺﾋﾟｰｻｲｽﾞ保存(保存用)
    ProgressBar1.Min = 0
    ProgressBar1.Max = AllCopySize

    'ComMng_Moduleｸﾗｽﾌﾟﾛﾊﾟﾃｨ設定
    cmm.ModuleSlot = SelectedSlot
    If ModuleType = CMDGA_TYPE Then
        cmm.RomType = ByteType
    Else
        cmm.RomType = WordType
    End If

End Sub


'--------------------------------------------------------
'
'   ROMｺﾋﾟｰ実行
'
'--------------------------------------------------------
Private Sub Form_Activate()

    Dim i As Integer

    '残像対策
    Me.Refresh
    
    'SPの運転を停止
    Call cmm.HaltSP
    If cmm.CmdState = CMD_ERR Then
        CopyMode = Error_Copy
        GoTo CopyEndCheck
    End If
    
    'ﾒｲﾝのSPをﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞに（ﾒｲﾝのSPがPFGAだった時のため）
    Call cmm.WriteSpInternalMemory(&H60B21A, &H80)
    If cmm.CmdState = CMD_ERR Then
        CopyMode = Error_Copy
        GoTo CopyEndCheck
    End If
    
    'ﾓｼﾞｭｰﾙﾀｲﾌﾟ毎にROM書込み処理
    Select Case ModuleType
    
        Case PFKP_TYPE, ETHER_TYPE
            'ROMｺﾋﾟｰﾙｰﾌﾟへ
            Call CopyLoop
            'MACｱﾄﾞﾚｽ書換え
            If ChangeMAC = ENABLE Then
                '---ｺﾋﾟｰ元ｽﾛｯﾄのﾏｯﾋﾟﾝｸﾞ解除
                cmm.ModuleSlot = SelectedSlot
                Call cmm.CancelMap
                cmm.ModuleSlot = TargetSlot
                '---ｾｸﾀ消去
                Call cmm.SectorEraseRom(MacWriteAddress)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---MACｱﾄﾞﾚｽ書込み
                Call cmm.WriteRom(MacWriteAddress, 10, WriteMacData())
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                End If
            End If
    
        Case CMDGA_TYPE
            'ROMは$20000番地～
            CopyAddress = CopyAddress + &H20000
            'ROMｺﾋﾟｰﾙｰﾌﾟへ
            Call CopyLoop
            
        Case DEVNET_TYPE
            For i = 0 To 1
                'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
                '---PFGA汎用ﾎﾟｰﾄ0ﾓｰﾄﾞﾚｼﾞｽﾀに1をﾗｲﾄ
                Call cmm.WriteMemoryByte(PORT0_MODE, &H1)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---PFGA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
                Call cmm.WriteMemoryByte(PORT0_CTLSTS, &H2)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '2周目はｺﾋﾟｰ先ｽﾛｯﾄのCPU
                cmm.ModuleSlot = TargetSlot
            Next
            'ROMｺﾋﾟｰﾙｰﾌﾟへ
            Call CopyLoop
        
        Case NETHER_TYPE
            For i = 0 To 1
                'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
                '---GA汎用ﾎﾟｰﾄ1ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT1_CTL, &H1)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT1_STS, &H2)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに0をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT1_STS, &H0)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '1秒ｳｪｲﾄ
                Sleep (2000)
                '2周目はｺﾋﾟｰ先ｽﾛｯﾄのCPU
                cmm.ModuleSlot = TargetSlot
            Next
            'ROMｺﾋﾟｰﾙｰﾌﾟへ
            Call CopyLoop
            'MACｱﾄﾞﾚｽ書換え
            If ChangeMAC = ENABLE Then
                '---ｺﾋﾟｰ元ｽﾛｯﾄのﾏｯﾋﾟﾝｸﾞ解除
                cmm.ModuleSlot = SelectedSlot
                Call cmm.CancelMap
                cmm.ModuleSlot = TargetSlot
                '---ｾｸﾀ消去
                Call cmm.SectorEraseRom(MacWriteAddress)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---MACｱﾄﾞﾚｽ書込み
                Call cmm.WriteRom(MacWriteAddress, 10, WriteMacData())
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                End If
            End If
    
        Case NSERIAL_TYPE
            For i = 0 To 1
                'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
                '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '2周目はｺﾋﾟｰ先ｽﾛｯﾄのCPU
                cmm.ModuleSlot = TargetSlot
            Next
            'ROMｺﾋﾟｰﾙｰﾌﾟへ
            Call CopyLoop
            
        Case RTOS_TYPE
            For i = 0 To 1
                'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
                '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
                Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
                If cmm.CmdState = CMD_ERR Then
                    CopyMode = Error_Copy
                    GoTo CopyEndCheck
                End If
                '2周目はｺﾋﾟｰ先ｽﾛｯﾄのCPU
                cmm.ModuleSlot = TargetSlot
            Next
            'ROMｺﾋﾟｰﾙｰﾌﾟへ
            Call CopyLoop
        
    End Select


CopyEndCheck:
    
    'ｺﾋﾟｰ元、ｺﾋﾟｰ先ﾓｼﾞｭｰﾙのﾏｯﾋﾟﾝｸﾞ解除
    cmm.ModuleSlot = SelectedSlot
    Call cmm.CancelMap
    cmm.ModuleSlot = TargetSlot
    Call cmm.CancelMap
    
    'ROMｺﾋﾟｰ終了
    Animation1.Stop
    Unload Me
    
    'ROMｺﾋﾟｰ完了確認
    Select Case CopyMode
        Case Normal_Copy
            MsgBox "ROMコピーを完了しました。", vbInformation, "SUCCESS"
            AddSysLog ("Slot" + Format(SelectedSlot, "00") + "⇒" + Format(TargetSlot, "00") + "：ROMコピー完了")
        Case Cancel_Copy
            MsgBox "ROMコピーを中断しました。", vbExclamation, "BREAK"
            AddSysLog ("Slot" + Format(SelectedSlot, "00") + "⇒" + Format(TargetSlot, "00") + "：ROMコピー中断")
        Case Error_Copy
            MsgBox "ROMコピーに失敗しました。", vbCritical, "ERROR"
            AddSysLog ("Slot" + Format(SelectedSlot, "00") + "⇒" + Format(TargetSlot, "00") + "：ROMコピー失敗")
    End Select

End Sub


'--------------------------------------------------------
'
'   ROMｺﾋﾟｰ中断
'
'--------------------------------------------------------
Private Sub Cancel_btn_Click()
    
    res = MsgBox("ROMコピーを中断しますか？", vbYesNo + vbExclamation, "REALLY?")
    Select Case res
        Case vbYes
            CopyMode = Cancel_Copy
        Case vbNo
            '何もしない
    End Select
End Sub


'--------------------------------------------------------
'
'   ROMｺﾋﾟｰﾙｰﾌﾟ
'
'--------------------------------------------------------
Private Sub CopyLoop()

    Do While CopySize > 0
        DoEvents
        'ROM読出し進捗表示
        Call DisplayRate(AllCopySize, CLng(CopySize))
        'ｺﾋﾟｰｷｬﾝｾﾙ
        If CopyMode = Cancel_Copy Then
            Exit Do
        End If
        'ﾓｼﾞｭｰﾙ内FlashRomｺﾋﾟｰ
        Call cmm.CopyRom(SelectedSlot, TargetSlot, CopyAddress, 1)
        '1MBit毎にｺﾋﾟｰ
        CopySize = CopySize - 1
        CopyAddress = CopyAddress + &H20000
        'ｴﾗｰ終了
        If cmm.CmdState = CMD_ERR Then
            CopyMode = Error_Copy
            Exit Do
        End If
    Loop

End Sub


'--------------------------------------------------------
'
'   ROMｺﾋﾟｰ進捗表示
'
'--------------------------------------------------------
Private Sub DisplayRate(AllSize As Long, CurrentSize As Long)

    Dim RateValue As Double
    
    RateValue = (CDbl(AllSize - CurrentSize) / CDbl(AllSize)) * 100
    Rate.Caption = Format(RateValue, "##0") + "%"
    Rate.Refresh
    ProgressBar1.Value = Int(AllSize - CurrentSize)
    ProgressBar1.Refresh
    If Main_Form.WindowState = vbMinimized Then
        Main_Form.Caption = Rate.Caption
    End If
    
End Sub

