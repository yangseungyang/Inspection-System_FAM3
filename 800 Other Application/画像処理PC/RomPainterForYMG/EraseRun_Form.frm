VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EraseRun_Form 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Erase Running"
   ClientHeight    =   1440
   ClientLeft      =   7665
   ClientTop       =   7395
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3015
   StartUpPosition =   2  '画面の中央
   Begin VB.Timer TimeoutTimer 
      Left            =   2400
      Top             =   720
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   57
      FullHeight      =   41
   End
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "中断"
      Height          =   420
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "EraseRun_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  EraseRun_Form                                       *
'*                                                      *
'*      作成日  ：2007/02/16                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM消去 実行画面                      *
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
'   ROM消去実行画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    '全消去は中断できない
    If EraseMode = Tip_Erase Then
        Cancel_btn.Enabled = False
    End If
    
    'ROM消去動画ｽﾀｰﾄ
    Animation1.Open App.Path + "\EraseAnimation.avi"
    Animation1.Play -1
    
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
'   ROM消去実行
'
'--------------------------------------------------------
Private Sub Form_Activate()

    '残像対策
    Me.Refresh

    'SPの運転を停止
    Call cmm.HaltSP
    If cmm.CmdState = CMD_ERR Then
        EraseMode = Error_Erase
        GoTo EraseEndCheck
    End If
    
    'ﾒｲﾝのSPをﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞに（ﾒｲﾝのSPがPFGAだった時のため）
    Call cmm.WriteSpInternalMemory(&H60B21A, &H80)
    If cmm.CmdState = CMD_ERR Then
        EraseMode = Error_Erase
        GoTo EraseEndCheck
    End If
    
    'ﾓｼﾞｭｰﾙﾀｲﾌﾟ毎のROM消去処理
    Select Case ModuleType
        
        Case PFKP_TYPE, ETHER_TYPE
            'ﾌｪｲﾙｱｸｾｽENABLE(PFGA用)
            Call cmm.WriteGAReg(SLAVE_CMD, &H100)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            'RDYを落とさないようにWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            'Call cmm.WriteGAInternalReg(SYS_CTL, &H1) 'PFGAの時は下位ｳｨﾝﾄﾞｳｱﾄﾞﾚｽで内部ﾚｼﾞｽﾀｱｸｾｽがないので×
            Call cmm.WriteMemoryByte(SYS_CTL_PFKP, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '全消去
            If EraseMode = Tip_Erase Then
                Call cmm.TipEraseRom(EraseAddress, 20)
                If cmm.CmdState = CMD_ERR Then
                    EraseMode = Error_Erase
                End If
            'ｾｸﾀ消去
            ElseIf EraseMode = Sector_Erase Then
                Call SectorEraseLoop
            End If
        
        Case CMDGA_TYPE
            'RDYを落とさないようにWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            Call cmm.WriteMemoryByte(ZIOSEL3, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '全消去
            If EraseMode = Tip_Erase Then
                Call cmm.TipEraseRom(&H20000, 30)
                If cmm.CmdState = CMD_ERR Then
                    EraseMode = Error_Erase
                End If
            'ｾｸﾀ消去
            ElseIf EraseMode = Sector_Erase Then
                Call SectorEraseLoop
            End If
        
        Case DEVNET_TYPE
            'ﾌｪｲﾙｱｸｾｽENABLE(PFGA用)
            Call cmm.WriteGAReg(SLAVE_CMD, &H100)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---PFGA汎用ﾎﾟｰﾄ0ﾓｰﾄﾞﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteMemoryByte(PORT0_MODE, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '---PFGA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteMemoryByte(PORT0_CTLSTS, &H2)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '全消去
            If EraseMode = Tip_Erase Then
                Call cmm.TipEraseRom(0, 20)
                If cmm.CmdState = CMD_ERR Then
                    EraseMode = Error_Erase
                End If
            'ｾｸﾀ消去
            ElseIf EraseMode = Sector_Erase Then
                Call SectorEraseLoop
            End If
        
        Case FLNET_TYPE
            'ﾓｼﾞｭｰﾙ内CPUをﾊﾞｽ解放状態に
            Call cmm.WriteMemoryWord(ROM_WRITE_ADDRESS, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                 GoTo EraseEndCheck
            End If
            'FPGAのKP_rom_windowﾚｼﾞｽﾀの初期値
            KpRomWindow = Int(EraseAddress \ 8192)
            Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                 GoTo EraseEndCheck
            End If
            '全消去
            If EraseMode = Tip_Erase Then
                Call cmm.TipEraseRom(0, 20)
                If cmm.CmdState = CMD_ERR Then
                    EraseMode = Error_Erase
                End If
            'ｾｸﾀ消去
            ElseIf EraseMode = Sector_Erase Then
                Call SectorEraseLoop
            End If
        
        Case NETHER_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ1ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに0をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_STS, &H0)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '1秒ｳｪｲﾄ
            Sleep (2000)
            '全消去
            If EraseMode = Tip_Erase Then
                Call cmm.TipEraseRom(0, 30)
                If cmm.CmdState = CMD_ERR Then
                    EraseMode = Error_Erase
                End If
            'ｾｸﾀ消去
            ElseIf EraseMode = Sector_Erase Then
                Call SectorEraseLoop
            End If
        
        Case NSERIAL_TYPE
            'KPGAのWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '全消去
            If EraseMode = Tip_Erase Then
                Call cmm.TipEraseRom(0, 10)
                If cmm.CmdState = CMD_ERR Then
                    EraseMode = Error_Erase
                End If
            'ｾｸﾀ消去
            ElseIf EraseMode = Sector_Erase Then
                Call SectorEraseLoop
            End If

        Case RTOS_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                EraseMode = Error_Erase
                GoTo EraseEndCheck
            End If
            '全消去
            If EraseMode = Tip_Erase Then
                Call cmm.TipEraseRom(0, 70) 'ﾀｲﾑｱｳﾄ約77秒
                If cmm.CmdState = CMD_ERR Then
                    EraseMode = Error_Erase
                    GoTo EraseEndCheck
                End If
            'ｾｸﾀ消去
            ElseIf EraseMode = Sector_Erase Then
                Call SectorEraseLoop
            End If
           
    End Select


EraseEndCheck:
    
    'ﾓｼﾞｭｰﾙのﾏｯﾋﾟﾝｸﾞ解除
    Call cmm.CancelMap
    
    'ROM消去動画終了
    Animation1.Close
    
    'ﾀｲﾑｱｳﾄﾀｲﾏ終了
    TimeoutTimer.Enabled = False
    
    '実行画面を閉じる
    Unload Me
    
    'ROM消去完了確認
'    If EraseMode = Cancel_Erase Then
'        '---消去中断
'        MsgBox "ROM消去を中断しました", vbExclamation, "BREAK"
'        AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM消去中断")
'    ElseIf EraseMode = Error_Erase Then
'        '---消去ﾒｿｯﾄﾞ失敗
'        MsgBox "ROM消去に失敗しました", vbCritical, "ERROR"
'        AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM消去失敗")
'    Else
'        '---消去正常終了確認
'        MsgBox "ROM消去を完了しました", vbInformation, "SUCCESS"
'        AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM消去完了")
'    End If
    
End Sub


'--------------------------------------------------------
'
'   ROM消去中断(ｾｸﾀ消去のみ有効)
'
'--------------------------------------------------------
Private Sub Cancel_btn_Click()

    res = MsgBox("ROM消去を中断しますか？", vbQuestion + vbYesNo, "REALLY?")
    Select Case res
        Case vbYes
            EraseMode = Cancel_Erase
        Case vbNo
            '何もしない
    End Select

End Sub


'--------------------------------------------------------
'
'   ROMｾｸﾀ消去ﾙｰﾌﾟ
'
'--------------------------------------------------------
Private Sub SectorEraseLoop()

    Do While EraseSize > 0
        DoEvents
        'ｺﾏﾝﾄﾞGAﾀｲﾌﾟのﾓｼﾞｭｰﾙは$20000番地以降をｾｸﾀ消去できない
        If ModuleType = CMDGA_TYPE Then
            If EraseAddress >= 131072 Then
                MsgBox "CMDGAタイプのモジュールは$20000番地以降セクタ消去できません。"
                EraseMode = Cancel_Erase
                Exit Do
            End If
        End If
        'ROM消去ﾒｿｯﾄﾞｺｰﾙ
        Call cmm.SectorEraseRom(EraseAddress)
        'ROM消去ﾒｿｯﾄﾞ失敗
        If cmm.CmdState = CMD_ERR Then
            EraseMode = Error_Erase
            Exit Do
        End If
        'ｾｸﾀ消去中断
        If EraseMode = Cancel_Erase Then Exit Do
        '8KByte単位でｾｸﾀ消去ｺﾏﾝﾄﾞ発行
        EraseAddress = EraseAddress + 8192
        EraseSize = EraseSize - 8
        'PFKP_TYPEのWindow切換え
        If ModuleType = FLNET_TYPE Then
            If (EraseAddress Mod 8192) = 0 Then
                'EraseAddress = 0
                KpRomWindow = KpRomWindow + 1
                Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            End If
        End If
    Loop
    
End Sub


'--------------------------------------------------------
'
'   ROM全消去(直接ROMｱｸｾｽ)のﾀｲﾑｱｳﾄ
'
'--------------------------------------------------------
Private Sub TimeoutTimer_Timer()
    
    EraseMode = Error_Erase
    TimeoutTimer.Enabled = False

End Sub

