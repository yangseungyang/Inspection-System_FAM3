VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VerifyRun_Form 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Verify Running"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4170
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Cancel_btn 
      Caption         =   "中断"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
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
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      _Version        =   393216
      FullWidth       =   49
      FullHeight      =   41
   End
   Begin VB.Label Job_lbl 
      Alignment       =   2  '中央揃え
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   405
      Width           =   1455
   End
   Begin VB.Label Rate 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "VerifyRun_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  VerifyRun_Form                                      *
'*                                                      *
'*      作成日  ：2007/05/09                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM照合 実行画面                      *
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
'   ROM照合実行画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()
        
    'ROM照合動画ｽﾀｰﾄ
    Animation1.Open WorkFolder + "\VerifyAnimation.avi"
    Animation1.Play -1
        
    '変数初期化
    McomDataNum = 512                              '一回の通信での照合ｻｲｽﾞ(MAX=512)
    VerifyErrorCount = 0                           '照合不一致数初期化
    If VerifyFirm = ENABLE Then
       AllVerifySize = VerifySize                  '全照合ｻｲｽﾞ保存
       ProgressBar1.Min = 0
       ProgressBar1.Max = AllVerifySize
       FilePointer = VerifyAddress + 1             '照合元ﾌｧｲﾙの読出し位置
       FileNo = FreeFile(0)                        '使用可能なﾌｧｲﾙ番号を取得
       Open SelectedFilePass For Binary As FileNo  '選択ﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
    End If
    
    '照合結果画面
    With VerifyResult_Form.MSFlexGrid1
        .Clear
        .Cols = 4
        .Rows = 11
        .ColWidth(0) = 300
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .Row = 0
        .Col = 1
        .Text = "ROMｱﾄﾞﾚｽ"
        .Col = 2
        .Text = "ﾌｧｲﾙﾃﾞｰﾀ"
        .Col = 3
        .Text = "ROMﾃﾞｰﾀ"
    End With
    'VerifyResult_Form.Caption = Hex(VerifyAddress) & "番地から" & (AllVerifySize / 1024) & "KByteの照合結果"
    VerifyResult_Form.Caption = "照合結果"
    VerifyResult_Form.Result_lbl = "照合中"
    
    'ComMng_Moduleｸﾗｽﾌﾟﾛﾊﾟﾃｨ設定
    cmm.ModuleSlot = SelectedSlot
    cmm.EndianMode = Endian
    If ModuleType = CMDGA_TYPE Then
        cmm.RomType = ByteType
    Else
        cmm.RomType = WordType
    End If

End Sub


'--------------------------------------------------------
'
'   ROM照合初期化
'
'--------------------------------------------------------
Private Sub Form_Activate()

    '残像対策
    Me.Refresh
    
    'SPの運転を停止
    Call cmm.HaltSP
    If cmm.CmdState = CMD_ERR Then
        VerifyMode = Error_Verify
        GoTo VerifyEndCheck
    End If
    
    'ﾒｲﾝのSPをﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞに（ﾒｲﾝのSPがPFGAだった時のため）
    Call cmm.WriteSpInternalMemory(&H60B21A, &H80)
    If cmm.CmdState = CMD_ERR Then
        VerifyMode = Error_Verify
        GoTo VerifyEndCheck
    End If
    
    'ﾓｼﾞｭｰﾙﾀｲﾌﾟ毎にROM照合処理
    Select Case ModuleType
        
        Case PFKP_TYPE, ETHER_TYPE
            'ﾌｪｲﾙｱｸｾｽENABLE(PFGA用)
            Call cmm.WriteGAReg(SLAVE_CMD, &H100)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'RDYを落とさないようにWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            Call cmm.WriteMemoryByte(SYS_CTL_PFKP, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'ROM照合ﾙｰﾌﾟへ
            Call VerifyLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
        
        Case CMDGA_TYPE
            'ｺﾏｹﾞｰの場合はROMは$20000番地～
            VerifyAddress = VerifyAddress + &H20000
            'CMDGAのWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            Call cmm.WriteMemoryByte(ZIOSEL3, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'ROM照合ﾙｰﾌﾟへ
            Call VerifyLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
        
        Case DEVNET_TYPE
            'ﾌｪｲﾙｱｸｾｽENABLE(PFGA用)
            Call cmm.WriteGAReg(SLAVE_CMD, &H100)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---PFGA汎用ﾎﾟｰﾄ0ﾓｰﾄﾞﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteMemoryByte(PORT0_MODE, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '---PFGA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteMemoryByte(PORT0_CTLSTS, &H2)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'F/W照合
            If VerifyFirm = ENABLE Then
                Job_lbl.Caption = "ﾌｧｰﾑｳｪｱ" & vbNewLine & "照合中"
                Call VerifyLoop                                     '照合ﾙｰﾌﾟ
                Close FileNo                                        'ﾌｧｲﾙを閉じる
                If (VerifyMode = Error_Verify) Or (VerifyMode = Over_Verify) Then
                    GoTo VerifyEndCheck
                End If
            End If
            'ﾃﾞﾊﾞｲｽﾈｯﾄ情報照合
            If VerifyDevInfo = ENABLE Then
                Job_lbl.Caption = "ﾃﾞﾊﾞｲｽﾈｯﾄ情報" & vbNewLine & "照合中"
                VerifyAddress = DevInfoAddress                      'ﾃﾞﾊﾞｲｽﾈｯﾄ情報先頭ｱﾄﾞﾚｽ
                VerifySize = DevInfoFileSize                        'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ｻｲｽﾞ
                AllVerifySize = DevInfoFileSize                     'ﾃﾞﾊﾞｲｽﾈｯﾄ情報照合ｻｲｽﾞ保存
                ProgressBar1.Min = 0
                ProgressBar1.Max = AllVerifySize
                FileNo = FreeFile(0)                                'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙのNo
                FilePointer = 1
                Open DevInfoFilePass For Binary As FileNo           'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
                Call VerifyLoop
                Close FileNo                                        'ﾌｧｲﾙを閉じる
                If (VerifyMode = Error_Verify) Or (VerifyMode = Over_Verify) Then
                    GoTo VerifyEndCheck
                End If
            End If
            'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報照合
            If VerifyCfgInfo = ENABLE Then
                Job_lbl.Caption = "ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報" & vbNewLine & "照合中"
                VerifyAddress = CfgInfoAddress                      'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報先頭ｱﾄﾞﾚｽ
                VerifySize = CfgInfoFileSize                        'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ｻｲｽﾞ
                AllVerifySize = CfgInfoFileSize                     'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報照合ｻｲｽﾞ保存
                ProgressBar1.Min = 0
                ProgressBar1.Max = AllVerifySize
                FileNo = FreeFile(0)                                'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙのNo
                FilePointer = 1
                Open CfgInfoFilePass For Binary As FileNo           'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
                Call VerifyLoop
                Close FileNo                                        'ﾌｧｲﾙを閉じる
                If (VerifyMode = Error_Verify) Or (VerifyMode = Over_Verify) Then
                    GoTo VerifyEndCheck
                End If
            End If
        
        Case FLNET_TYPE
            'ﾓｼﾞｭｰﾙ内CPUをﾊﾞｽ解放状態に
            Call cmm.WriteMemoryWord(ROM_WRITE_ADDRESS, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'FPGAのKP_rom_windowﾚｼﾞｽﾀの初期値
            KpRomWindow = Int(VerifyAddress \ 8192)
            Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            VerifyAddress = VerifyAddress Mod 8192
            'ROM照合ﾙｰﾌﾟへ
            Call VerifyLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
            
        Case NETHER_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ1ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに0をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_STS, &H0)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '1秒ｳｪｲﾄ
            Sleep (1000)
            'ROM照合ﾙｰﾌﾟへ
            Call VerifyLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
        
        Case NSERIAL_TYPE
            'KPGAのWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            'ROM照合ﾙｰﾌﾟへ
            Call VerifyLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
        
        Case RTOS_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                VerifyMode = Error_Verify
                GoTo VerifyEndCheck
            End If
            '基本ﾌﾞｰﾄ照合
            If VerifyFirm = ENABLE Then
                Job_lbl.Caption = "基本ﾌﾞｰﾄ" & vbNewLine & "照合中"
                Call VerifyLoop                                 '基本ﾌﾞｰﾄ照合
                Close FileNo                                    'ﾌｧｲﾙを閉じる
                If (VerifyMode = Error_Verify) Or (VerifyMode = Over_Verify) Then
                    GoTo VerifyEndCheck
                End If
            End If
            'OSﾌﾞｰﾄ照合
            If VerifySystem = ENABLE Then
                Job_lbl.Caption = "OSﾌﾞｰﾄ" & vbNewLine & "照合中"
                VerifyAddress = SystemAddress                   'OSﾌﾞｰﾄ先頭ｱﾄﾞﾚｽ
                VerifySize = SystemFileSize                     'OSﾌﾞｰﾄｻｲｽﾞ
                AllVerifySize = SystemFileSize                  'OSﾌﾞｰﾄ照合ｻｲｽﾞ保存
                ProgressBar1.Min = 0
                ProgressBar1.Max = AllVerifySize
                FileNo = FreeFile(0)                            'OSﾌﾞｰﾄﾌｧｲﾙのNo
                FilePointer = 1
                Open SystemFilePass For Binary As FileNo        'OSﾌﾞｰﾄﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
                Call VerifyLoop                                 'OSﾌﾞｰﾄ照合
                Close FileNo                                    'ﾌｧｲﾙを閉じる
                If (VerifyMode = Error_Verify) Or (VerifyMode = Over_Verify) Then
                    GoTo VerifyEndCheck
                End If
            End If
    End Select
    
    
VerifyEndCheck:
    
    'ﾓｼﾞｭｰﾙのﾏｯﾋﾟﾝｸﾞ解除
    Call cmm.CancelMap
    
    'ROM照合動画終了
    Animation1.Stop
    Unload Me
    
    'ROM照合終了確認
    Select Case VerifyMode
        Case Normal_Verify
            If VerifyErrorCount = 0 Then
                MsgBox "一致しました", vbInformation, "SUCCESS"
            Else
                VerifyResult_Form.Result_lbl.Caption = "不一致箇所が見つかりました"
                VerifyResult_Form.Show vbModal
            End If
            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM照合完了")
        Case Cancel_Verify
            MsgBox "ROM照合を中断しました", vbExclamation, "BREAK"
            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM照合中断")
        Case Over_Verify
            MsgBox "不一致箇所が10箇所以上見つかりました", vbExclamation, "BREAK"
            VerifyResult_Form.Result_lbl.Caption = "不一致箇所が見つかりました"
            VerifyResult_Form.Show vbModal
            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM照合中断")
        Case Error_Verify
            MsgBox "ROM照合に失敗しました", vbCritical, "ERROR"
            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM照合失敗")
    End Select
    
    'ROM照合結果画面を閉じる
    Unload VerifyResult_Form
    
End Sub


'--------------------------------------------------------
'
'   ｷｬﾝｾﾙﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Cancel_btn_Click()

    res = MsgBox("ROM照合を中断しますか？", vbQuestion + vbYesNo, "BREAK")
    Select Case res
        Case vbYes
            VerifyMode = Cancel_Verify
        Case vbNo
            '何もしない
    End Select

End Sub


'--------------------------------------------------------
'
'   ROM照合ﾙｰﾌﾟ
'
'--------------------------------------------------------
Private Sub VerifyLoop()
 
    Dim i As Integer
    Dim LoopNum As Integer
 
    Dim r_RomData() As Byte             '照合する読み出したﾌｧｰﾑのﾃﾞｰﾀ
    ReDim r_RomData(McomDataNum - 1)
    
    Dim v_FileData() As Byte            '照合するﾌｧｲﾙのﾃﾞｰﾀ
    ReDim v_FileData(McomDataNum - 1)
    
    LoopNum = 0
    
    Do While VerifySize > 0
        DoEvents
        'ROM照合進捗表示
        Call DisplayRate(AllVerifySize, VerifySize)
        'ﾓｼﾞｭｰﾙ内FlashRom読出し
        Call cmm.ReadRom(VerifyAddress, McomDataNum, r_RomData())
        'ｴﾗｰ終了
        If cmm.CmdState = CMD_ERR Then
            VerifyMode = Error_Verify
            Exit Do
        End If
        '照合ｷｬﾝｾﾙ
        If VerifyMode = Cancel_Verify Then
            Exit Do
        End If
        '照合元ﾌｧｲﾙのﾃﾞｰﾀ取得
        Get FileNo, FilePointer, v_FileData()
        'ﾃﾞｰﾀ比較
        For i = 0 To McomDataNum - 1
            If ModuleType <> FLNET_TYPE And VerifyMAC = DISABLE And _
                MacWrittenAddress <= VerifyAddress + i + LoopNum * 512 And _
                    VerifyAddress + i + LoopNum * 512 <= MacWrittenAddress + 9 And _
                    MacWrittenAddress <> 0 Then
                '(MACｱﾄﾞﾚｽ非照合範囲内の場合は何もしない)
            ElseIf ModuleType = FLNET_TYPE And VerifyMAC = DISABLE And _
                MacWrittenAddress <= CLng(KpRomWindow) * 8192 + i + LoopNum * 512 And _
                    CLng(KpRomWindow) * 8192 + i + LoopNum * 512 <= MacWrittenAddress + 9 And _
                    MacWrittenAddress <> 0 Then
                '(MACｱﾄﾞﾚｽ非照合範囲内の場合は何もしない)
            ElseIf r_RomData(i) <> v_FileData(i) Then
                VerifyErrorCount = VerifyErrorCount + 1
                '照合結果画面に反映
                With VerifyResult_Form.MSFlexGrid1
                    .Row = VerifyErrorCount
                    .Col = 0
                    .Text = VerifyErrorCount
                    .Col = 1
                    If ModuleType = FLNET_TYPE Then
                        .Text = Hex(CLng(KpRomWindow) * 8192 + i + LoopNum * 512)
                    ElseIf ModuleType = CMDGA_TYPE Then
                        .Text = Hex(VerifyAddress - 131072 + i)
                    Else
                        .Text = Hex(VerifyAddress + i)
                    End If
                    .Col = 2
                    .Text = Right$("0" & Hex(v_FileData(i)), 2)
                    .Col = 3
                    .Text = Right$("0" & Hex(r_RomData(i)), 2)
                End With
                '10個ｴﾗｰで照合終了
                If VerifyErrorCount = 10 Then
                    VerifyMode = Over_Verify
                    Exit Do
                End If
            End If
        Next
        'McomDataNum毎に読出し
        VerifyAddress = VerifyAddress + McomDataNum
        VerifySize = VerifySize - McomDataNum
        FilePointer = FilePointer + McomDataNum
        'PFKP_TYPEのWindow切換え
        If ModuleType = FLNET_TYPE Then
            LoopNum = LoopNum + 1
            If (VerifyAddress Mod 8192) = 0 Then
                VerifyAddress = 0
                KpRomWindow = KpRomWindow + 1
                LoopNum = 0
                Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            End If
        End If
    Loop
    
End Sub


'--------------------------------------------------------
'
'   ROM照合進捗表示
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
