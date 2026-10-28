VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ReadRun_Form 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "Rom Read Running"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4020
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4020
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
   Begin VB.Label Rate 
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "ReadRun_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  ReadRun_Form                                        *
'*                                                      *
'*      作成日  ：2007/03/02                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM読出し 実行画面                    *
'*                                                      *
'********************************************************


'Option Explicit

'---------------------------------------------
'外部API関数使用宣言
'---------------------------------------------
'---Sleep関数
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------
'ｵﾌﾞｼﾞｪｸﾄ変数
'---------------------------------------------
Dim tmpFile As File
Dim tmpfso As New FileSystemObject


'--------------------------------------------------------
'
'   ROM読出し実行画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()

    'ROM読出し動画ｽﾀｰﾄ
    Animation1.Open WorkFolder + "\ReadAnimation.avi"
    Animation1.Play -1
    
    McomDataNum = 512                           'ﾒｿｯﾄﾞ1回の読出しｻｲｽﾞ(max992)
    
    AllReadSize = ReadSize                      '選択ﾌｧｲﾙのｻｲｽﾞ保存(保存用)
    ProgressBar1.Min = 0
    ProgressBar1.Max = AllReadSize
    
    '読出し先ﾌｧｲﾙ
    If ReadMode = Overwrite_Read Then           '上書きﾓｰﾄﾞの際は元のﾌｧｲﾙを削除する
        Set tmpFile = tmpfso.GetFile(ReadFilePass)
        tmpFile.Delete
    End If
    FileNo = FreeFile(0)                        '使用可能なﾌｧｲﾙ番号を取得
    FilePointer = 1                             '書込み元ﾌｧｲﾙの読出し位置
    Open ReadFilePass For Binary As FileNo      '選択ﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
    
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
'   ROM読出し実行
'
'--------------------------------------------------------
Private Sub Form_Activate()

    Dim b_WriteData(0) As Byte
    Dim i_WriteData(0) As Integer

    '残像対策
    Me.Refresh
    
    'SPの運転を停止
    Call cmm.HaltSP
    If cmm.CmdState = CMD_ERR Then
        ReadMode = Error_Read
        GoTo ReadEndCheck
    End If
    
    'ﾒｲﾝのSPをﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞに（ﾒｲﾝのSPがPFGAだった時のため）
    Call cmm.WriteSpInternalMemory(&H60B21A, &H80)
    If cmm.CmdState = CMD_ERR Then
        ReadMode = Error_Read
        GoTo ReadEndCheck
    End If
    
    'ﾓｼﾞｭｰﾙﾀｲﾌﾟ毎にROM読出し処理
    Select Case ModuleType
    
        Case PFKP_TYPE, ETHER_TYPE
            'ﾌｪｲﾙｱｸｾｽENABLE(PFGA用)
            Call cmm.WriteGAReg(SLAVE_CMD, &H100)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'RDYを落とさないようにWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            Call cmm.WriteMemoryByte(SYS_CTL_PFKP, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ROM読出しﾙｰﾌﾟへ
            Call ReadLoop
            'MACｱﾄﾞﾚｽ初期化(ｵﾌﾟｼｮﾝ画面でENABLEにする)
            If InitMacAddress = ENABLE Then
                '---ﾌｧｲﾙにﾃﾞﾌｫﾙﾄMACｱﾄﾞﾚｽ書込み
                Put FileNo, MacReadAddress, DefaultMAC()
            End If
    
        Case CMDGA_TYPE
            'ROMは$20000番地～
            ReadAddress = ReadAddress + &H20000
            'RDYを落とさないようにWDT停止
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            Call cmm.WriteMemoryByte(ZIOSEL3, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ROM読出しﾙｰﾌﾟへ
            Call ReadLoop
    
        Case DEVNET_TYPE
            'ﾌｪｲﾙｱｸｾｽENABLE(PFGA用)
            Call cmm.WriteGAReg(SLAVE_CMD, &H100)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---PFGA汎用ﾎﾟｰﾄ0ﾓｰﾄﾞﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteMemoryByte(PORT0_MODE, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '---PFGA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteMemoryByte(PORT0_CTLSTS, &H2)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ROM読出しﾙｰﾌﾟへ
            Call ReadLoop
    
        Case FLNET_TYPE
            'ﾓｼﾞｭｰﾙ内CPUをﾊﾞｽ解放状態に
            Call cmm.WriteMemoryWord(ROM_WRITE_ADDRESS, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'FPGAのKP_rom_windowﾚｼﾞｽﾀの初期値
            KpRomWindow = ReadAddress \ 8192
            Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            ReadAddress = ReadAddress Mod 8192
            'ROM読出しﾙｰﾌﾟへ
            Call ReadLoop
            'MACｱﾄﾞﾚｽ初期化
            If InitMacAddress = ENABLE Then
                '---ﾌｧｲﾙにﾃﾞﾌｫﾙﾄMACｱﾄﾞﾚｽ書込み
                Put FileNo, MacReadAddress, DefaultMAC()
            End If
            
        Case NETHER_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ1ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀに0をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT1_STS, &H0)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '1秒ｳｪｲﾄ
            Sleep (1000)
            'ROM読出しﾙｰﾌﾟへ
            Call ReadLoop
            'MACｱﾄﾞﾚｽ初期化
            If InitMacAddress = ENABLE Then
                '---ﾌｧｲﾙにﾃﾞﾌｫﾙﾄMACｱﾄﾞﾚｽ書込み
                Put FileNo, MacReadAddress, DefaultMAC()
            End If
    
        Case NSERIAL_TYPE
            'RDYを落とさないようにWDT停止(止めなきゃ駄目なのがいるかもしれないので)
            Call cmm.WriteGAReg(WDT_CTL, &H0)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ROM読出しﾙｰﾌﾟへ
            Call ReadLoop
        
        Case RTOS_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀに1をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀに2をﾗｲﾄ
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                ReadMode = Error_Read
                GoTo ReadEndCheck
            End If
            'ROM読出しﾙｰﾌﾟへ
            Call ReadLoop
        
    End Select

ReadEndCheck:
    
    'ﾓｼﾞｭｰﾙのﾏｯﾋﾟﾝｸﾞ解除
    Call cmm.CancelMap
    
    'ﾌｧｲﾙを閉じる
    Close FileNo
    
    'ROM読出し終了
    Animation1.Stop
    Unload Me
    
    'ROM読出し終了確認
    Select Case ReadMode
        Case Normal_Read, Overwrite_Read
            MsgBox "ROM読出しを完了しました", vbInformation, "SUCCESS"
            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM読出し完了")
        Case Cancel_Read
            MsgBox "ROM読出しを中断しました", vbExclamation, "SUCCESS"
            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM読出し中断")
        Case Error_Read
            MsgBox "ROM読出しに失敗しました", vbCritical, "SUCCESS"
            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM読出し失敗")
    End Select
    

End Sub


'--------------------------------------------------------
'
'   ｷｬﾝｾﾙﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Cancel_btn_Click()

    res = MsgBox("ROM読出しを中断しますか？", vbQuestion + vbYesNo, "REALLY?")
    Select Case res
        Case vbYes
            ReadMode = Cancel_Read
        Case vbNo
            '何もしない
    End Select

End Sub


'--------------------------------------------------------
'
'   ROM読出しﾙｰﾌﾟ
'
'--------------------------------------------------------
Private Sub ReadLoop()

    Dim i_WriteData(0) As Integer

    Dim ReadFirmData() As Byte
    ReDim ReadFirmData(McomDataNum - 1)

    Do While ReadSize > 0
        DoEvents
        'ROM読出し進捗表示
        Call DisplayRate(AllReadSize, ReadSize)
        'ﾓｼﾞｭｰﾙ内FlashRom読出し
        Call cmm.ReadRom(ReadAddress, McomDataNum, ReadFirmData())
        'ﾌｧｲﾙに書込み
        Put FileNo, FilePointer, ReadFirmData()
        'ReadSize毎に読出し
        ReadSize = ReadSize - McomDataNum
        FilePointer = FilePointer + McomDataNum
        ReadAddress = ReadAddress + McomDataNum
        'PFKP_TYPEのWindow切換え
        If ModuleType = FLNET_TYPE Then
            If (ReadAddress Mod 8192) = 0 Then
                ReadAddress = 0
                'Call cmm.WriteModuleMemoryWord(ROM_WRITE_ADDRESS, &H1)
                KpRomWindow = KpRomWindow + 1
                Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            End If
        End If
        'ｴﾗｰ終了
        If cmm.CmdState = CMD_ERR Then
            ReadMode = Error_Read
            Exit Do
        End If
        '読出しｷｬﾝｾﾙ
        If ReadMode = Cancel_Read Then
            Exit Do
        End If
    Loop

End Sub


'--------------------------------------------------------
'
'   ROM読出し進捗表示
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



'    saveMcomDataNum = McomDataNum
'
'        Case NETSP_TYPE
'            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
'            Call cmm.WriteMemoryWord(TRISTATE_REG, &H8000)
'            If cmm.CmdState = CMD_ERR Then
'                ReadMode = Error_Read
'                 GoTo ReadEndCheck
'            End If
'            'MXA_Uﾚｼﾞｽﾀにｱｸｾｽするﾌﾗｯｼｭのｱﾄﾞﾚｽの上位ﾋﾞｯﾄを設定
'            MXA_U_Value = (ReadAddress And &HE00000) / 65536
'            Call cmm.WriteMemoryWord(MXA_U_REG, MXA_U_Value)
'            'ROM読出しﾙｰﾌﾟへ
'            ReadAddress = ReadAddress Mod &H200000
'            Call ReadLoop


'        'NETSP_TYPEのWindow切換え
'        If ModuleType = NETSP_TYPE Then
'            'MXA_Uﾚｼﾞｽﾀにｱｸｾｽするﾌﾗｯｼｭのｱﾄﾞﾚｽの上位ﾋﾞｯﾄを設定
'            If ReadAddress >= &H200000 Then
'                MXA_U_Value = MXA_U_Value + &H20
'                Call cmm.WriteMemoryWord(MXA_U_REG, MXA_U_Value)
'                ReadAddress = ReadAddress Mod &H200000
'            End If
'            '200000番地の境界で最後の一回
'            If (&H200000 - ReadAddress) < McomDataNum Then
'                McomDataNum = &H200000 - ReadAddress
'            Else
'                McomDataNum = saveMcomDataNum
'            End If
'            '帳尻合わせ
'            If ReadSize < McomDataNum Then
'                McomDataNum = ReadSize
'            End If
'            ReDim ReadFirmData(McomDataNum - 1)
'        End If

