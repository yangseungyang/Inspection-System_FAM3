VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form WriteRun_Form 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "ROM Write Running"
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
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
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
   Begin VB.Label Job_lbl 
      Alignment       =   2  '中央揃え
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   285
      Width           =   1575
   End
   Begin VB.Label Rate 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "WriteRun_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                                                      *
'*  WriteRun_Form                                       *
'*                                                      *
'*      作成日  ：2007/02/23                            *
'*      作成者  ：澤原                                  *
'*      ｺﾒﾝﾄ    ：ROM書込み 実行画面                    *
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
'   ROM書込み実行画面の初期化
'
'--------------------------------------------------------
Private Sub Form_Load()
    
    'ROM書込み動画ｽﾀｰﾄ
    Animation1.Open App.Path + "\WriteAnimation.avi"
    Animation1.Play -1
    
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
'   ROM書込み実行
'
'--------------------------------------------------------
Private Sub Form_Activate()

    Dim s_ReadModuleName(9) As Byte
    Dim tmpWriteFilePass As String

    '残像対策
    Me.Refresh
    
    'SPの運転を停止
    Call cmm.HaltSP
    If cmm.CmdState = CMD_ERR Then
        WriteMode = Error_Write
        GoTo WriteEndCheck
    End If
    
    'ﾒｲﾝのSPをﾃﾞﾊﾞｯｸﾞﾓｰﾄﾞに（ﾒｲﾝのSPがPFGAだった時のため）
    Call cmm.WriteSpInternalMemory(&H60B21A, &H80)
    If cmm.CmdState = CMD_ERR Then
        WriteMode = Error_Write
        GoTo WriteEndCheck
    End If
    
    'ﾌｧｰﾑ書込み有りの場合ﾌｧｲﾙを開く
    If WriteFirm = ENABLE Then
        If ModuleType = FLNET_TYPE Then
            McomDataNum = 512                       'ﾒｿｯﾄﾞ1回の書込みｻｲｽﾞ(FLnetのみ)
        Else
            McomDataNum = 992                       'ﾒｿｯﾄﾞ1回の書込みｻｲｽﾞ(MAX)
        End If
        AllWriteSize = WriteSize                    '選択ﾌｧｲﾙのｻｲｽﾞ保存(保存用)
        ProgressBar1.Min = 0
        ProgressBar1.Max = AllWriteSize
        FileNo = FreeFile(0)                        '使用可能なﾌｧｲﾙ番号を取得
        FilePointer = 1                             '書込み元ﾌｧｲﾙの読出し位置
        
        Open SelectedFilePass For Binary As FileNo  '選択ﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
    
    End If
 
    
    'ﾓｼﾞｭｰﾙﾀｲﾌﾟ毎にROM書込み処理
    Select Case ModuleType
        
        Case PFKP_TYPE, ETHER_TYPE
            'ROM書込みﾙｰﾌﾟへ
            Call WriteLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
    
        Case CMDGA_TYPE
            'ROMは$20000番地～
            WriteAddress = Writeddress + &H20000
            'ROM書込みﾙｰﾌﾟへ
            Call WriteLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
    
        Case DEVNET_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---PFGA汎用ﾎﾟｰﾄ0ﾓｰﾄﾞﾚｼﾞｽﾀ=1
            Call cmm.WriteMemoryByte(PORT0_MODE, &H1)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '---PFGA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙｽﾃｰﾀｽﾚｼﾞｽﾀ=2
            Call cmm.WriteMemoryByte(PORT0_CTLSTS, &H2)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            'F/W書込み
            If WriteFirm = ENABLE Then
                Job_lbl.Caption = "ﾌｧｰﾑｳｪｱ" & vbNewLine & "書込み中"
                Call WriteLoop                              'F/W書込み
                Close FileNo                                'ﾌｧｲﾙを閉じる
                If WriteMode = Error_Write Then
                    GoTo WriteEndCheck
                End If
            End If
            'ﾃﾞﾊﾞｲｽﾈｯﾄ情報書込み
            If WriteDevInfo = ENABLE Then
                Job_lbl.Caption = "ﾃﾞﾊﾞｲｽﾈｯﾄ情報" & vbNewLine & "書込み中"
                McomDataNum = 512
                WriteAddress = DevInfoAddress               'ﾃﾞﾊﾞｲｽﾈｯﾄ情報書込み開始ｱﾄﾞﾚｽ
                WriteSize = DevInfoFileSize                 'ﾃﾞﾊﾞｲｽﾈｯﾄ情報書込みｻｲｽﾞ
                AllWriteSize = DevInfoFileSize              'ﾃﾞﾊﾞｲｽﾈｯﾄ情報書込みｻｲｽﾞ保存
                ProgressBar1.Min = 0
                ProgressBar1.Max = AllWriteSize
                FileNo = FreeFile(0)                        'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙのNo
                FilePointer = 1
                Open DevInfoFilePass For Binary As FileNo   'ﾃﾞﾊﾞｲｽﾈｯﾄ情報ﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
                Call WriteLoop                              'ﾃﾞﾊﾞｲｽﾈｯﾄ情報書込み
                Close FileNo                                'ﾌｧｲﾙを閉じる
                If WriteMode = Error_Write Then
                    GoTo WriteEndCheck
                End If
            End If
            'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報書込み
            If WriteCfgInfo = ENABLE Then
                Job_lbl.Caption = "ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報" & vbNewLine & "書込み中"
                McomDataNum = 512
                WriteAddress = CfgInfoAddress               'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報書込み開始ｱﾄﾞﾚｽ
                WriteSize = CfgInfoFileSize                 'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報書込みｻｲｽﾞ
                AllWriteSize = CfgInfoFileSize              'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報書込みｻｲｽﾞ保存
                ProgressBar1.Min = 0
                ProgressBar1.Max = AllWriteSize
                FileNo = FreeFile(0)                        'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙのNo
                FilePointer = 1
                Open CfgInfoFilePass For Binary As FileNo   'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報ﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
                Call WriteLoop                              'ｺﾝﾌｨｸﾞﾚｰｼｮﾝ情報書込み
                Close FileNo                                'ﾌｧｲﾙを閉じる
                If WriteMode = Error_Write Then
                    GoTo WriteEndCheck
                End If
            End If
        
        Case FLNET_TYPE
            'ﾓｼﾞｭｰﾙ内CPUをﾊﾞｽ解放状態に
            Call cmm.WriteMemoryWord(ROM_WRITE_ADDRESS, &H1)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            'FPGAのKP_rom_windowﾚｼﾞｽﾀの初期値
            KpRomWindow = Int(WriteAddress \ 8192)
            Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            'ROM書込みﾙｰﾌﾟへ
            Call WriteLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
        
        Case NETHER_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ1ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀ=1
            Call cmm.WriteGAInternalReg(PORT1_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀ=2
            Call cmm.WriteGAInternalReg(PORT1_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀ=1
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀ=2
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ1ｽﾃｰﾀｽﾚｼﾞｽﾀ=0
            Call cmm.WriteGAInternalReg(PORT1_STS, &H0)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '1秒ｳｪｲﾄ
            Sleep (1000)
            'ROM書込みﾙｰﾌﾟへ
            Call WriteLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
        
        Case NSERIAL_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀ=1
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀ=2
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            'ROM書込みﾙｰﾌﾟへ
            Call WriteLoop
            'ﾌｧｲﾙを閉じる
            Close FileNo
        
        Case RTOS_TYPE
            'ﾓｼﾞｭｰﾙ内CPUの動作をHALT
            '---GA汎用ﾎﾟｰﾄ0ｺﾝﾄﾛｰﾙﾚｼﾞｽﾀ=1
            Call cmm.WriteGAInternalReg(PORT0_CTL, &H1)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '---GA汎用ﾎﾟｰﾄ0ｽﾃｰﾀｽﾚｼﾞｽﾀ=2
            Call cmm.WriteGAInternalReg(PORT0_STS, &H2)
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            '基本ﾌﾞｰﾄ書込み
            If WriteFirm = ENABLE Then
                Job_lbl.Caption = "基本ﾌﾞｰﾄ" & vbNewLine & "書込み中"
                Call WriteLoop          '基本ﾌﾞｰﾄ書込み
                Close FileNo            'ﾌｧｲﾙを閉じる
            End If
            If cmm.CmdState = CMD_ERR Then
                WriteMode = Error_Write
                GoTo WriteEndCheck
            End If
            'OSﾌﾞｰﾄ書込み
            If WriteSystem = ENABLE Then
                Job_lbl.Caption = "OSﾌﾞｰﾄ" & vbNewLine & "書込み中"
                McomDataNum = 992
                WriteAddress = SystemAddress              'OSﾌﾞｰﾄ先頭ｱﾄﾞﾚｽ
                WriteSize = SystemFileSize                'OSﾌﾞｰﾄﾌｧｲﾙのｻｲｽﾞ
                AllWriteSize = SystemFileSize
                ProgressBar1.Min = 0
                ProgressBar1.Max = AllWriteSize
                FileNo = FreeFile(0)                      '使用可能なﾌｧｲﾙ番号を取得
                FilePointer = 1                           'OSﾌﾞｰﾄﾌｧｲﾙの読出し位置
                Open SystemFilePass For Binary As FileNo  'OSﾌﾞｰﾄﾌｧｲﾙをﾊﾞｲﾅﾘﾓｰﾄﾞで開く
                Call WriteLoop                            'OSﾌﾞｰﾄ書込み
                Close FileNo                              'ﾌｧｲﾙを閉じる
                If cmm.CmdState = CMD_ERR Then
                    WriteMode = Error_Write
                    GoTo WriteEndCheck
                End If
            End If
            'ﾓｼﾞｭｰﾙ形名書込み
            If WriteName = ENABLE Then
                ProgressBar1.Visible = False
                Rate.Caption = ""
                Label1.Caption = "ﾓｼﾞｭｰﾙ形名" & vbNewLine & "書込み中"
                '---ｾｸﾀ消去(ｾｸﾀ6)
                Call cmm.SectorEraseRom(RTOS_MODULENAME_ADDRESS)
                If cmm.CmdState = CMD_ERR Then
                    WriteMode = Error_Write
                    GoTo WriteEndCheck
                End If
                '---ﾓｼﾞｭｰﾙ形名書込み
                Call cmm.WriteRom(RTOS_MODULENAME_ADDRESS, 10, WriteModuleNameData())
                If cmm.CmdState = CMD_ERR Then
                    WriteMode = Error_Write
                    GoTo WriteEndCheck
                End If
            End If
            'MACｱﾄﾞﾚｽ書込み
            If WriteMAC = ENABLE Then
                Label1.Caption = "MACｱﾄﾞﾚｽ書込み中"
                '---ｾｸﾀ消去(ｾｸﾀ7)
                Call cmm.SectorEraseRom(RTOS_MACADDRESS_ADDRESS)
                If cmm.CmdState = CMD_ERR Then
                    WriteMode = Error_Write
                    GoTo WriteEndCheck
                End If
                '---MACｱﾄﾞﾚｽ書込み
                Call cmm.WriteRom(RTOS_MACADDRESS_ADDRESS, 6, WriteMacData())
                If cmm.CmdState = CMD_ERR Then
                    WriteMode = Error_Write
                    GoTo WriteEndCheck
                End If
            End If

    End Select


WriteEndCheck:

    'ﾓｼﾞｭｰﾙのﾏｯﾋﾟﾝｸﾞ解除
    Call cmm.CancelMap

    'MAC付書込みﾌｧｲﾙがあったら削除する
    tmpWriteFilePass = WorkFolder + "\tmpWriteFileWithMac.bin"
    If tmpfso.FileExists(tmpWriteFilePass) Then
        Set tmpFile = tmpfso.GetFile(tmpWriteFilePass)
        tmpFile.Delete
        'ﾌｧｲﾙのﾊﾟｽを元に戻す
        SelectedFilePass = SaveFilePass
    End If
            
    'ROM書込み終了
    Animation1.Stop
    Unload Me
    
'    'ROM書込み完了確認
'    Select Case WriteMode
'        Case Normal_Write
'            MsgBox "ROM書込みを完了しました", vbInformation, "SUCCESS"
'            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM書込み完了")
'        Case Cancel_Write
'            MsgBox "ROM書込みを中断しました", vbExclamation, "BREAK"
'            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM書込み中断")
'        Case Error_Write
'            MsgBox "ROM書込みに失敗しました", vbCritical, "ERROR"
'            AddSysLog ("Slot" + Format(cmm.ModuleSlot, "00") + "：ROM書込み失敗")
'    End Select

End Sub


'--------------------------------------------------------
'
'   ｷｬﾝｾﾙﾎﾞﾀﾝ
'
'--------------------------------------------------------
Private Sub Cancel_btn_Click()
   
    res = MsgBox("ROM書込みを中断しますか？", vbQuestion + vbYesNo, "REALLY?")
    Select Case res
        Case vbYes
            WriteMode = Cancel_Write
        Case vbNo
            '何もしない
    End Select
    
End Sub


'--------------------------------------------------------
'
'   ROM書込みﾙｰﾌﾟ
'
'--------------------------------------------------------
Private Sub WriteLoop()

    Dim WriteFirmData() As Byte     '書込むﾃﾞｰﾀ
    ReDim WriteFirmData(McomDataNum - 1)

    Dim BorderLine As Long
    Dim Remain As Long
    Dim svMcomDataNum As Integer
    svMcomDataNum = McomDataNum     'ﾒｿｯﾄﾞ一回の通信ﾃﾞｰﾀ量を保存しておく

    Do While WriteSize > 0
        DoEvents
        'ROM書込み進捗表示
        Call DisplayRate(AllWriteSize, WriteSize)
        '書込みﾌｧｲﾙのﾃﾞｰﾀ取得
        Get FileNo, FilePointer, WriteFirmData()
        'ROM書込みﾒｿｯﾄﾞ呼び出し
        Call cmm.WriteRom(WriteAddress, McomDataNum, WriteFirmData())
        'McomDataNum毎に書込み
        WriteAddress = WriteAddress + McomDataNum
        WriteSize = WriteSize - McomDataNum
        FilePointer = FilePointer + McomDataNum
        '64KB以上F/W用、64KBの境界
        BorderLine = WriteAddress Mod 65536
        If BorderLine > (65536 - McomDataNum) Then
            Remain = 65536 - BorderLine
            If Remain < 0 Then
                Remain = Remain + 65536
            End If
            McomDataNum = CInt(Remain)
        Else
            McomDataNum = svMcomDataNum
        End If
        '最後の一回の書込みｻｲｽﾞ
        If WriteSize < McomDataNum Then
            McomDataNum = WriteSize
        End If
        'PFKP_TYPEのWindow切換え
        If ModuleType = FLNET_TYPE Then
            If (WriteAddress Mod 8192) = 0 Then
                WriteAddress = 0
                'Call cmm.WriteModuleMemoryWord(ROM_WRITE_ADDRESS, &H1)
                KpRomWindow = KpRomWindow + 1
                Call cmm.WriteMemoryWord(ROM_WINDOW_ADDRESS, KpRomWindow)
            End If
        End If
        'ｴﾗｰ終了
        If cmm.CmdState = CMD_ERR Then
            WriteMode = Error_Write
            Exit Do
        End If
        '書込みｷｬﾝｾﾙ
        If WriteMode = Cancel_Write Then
            Exit Do
        End If
    Loop

End Sub


'--------------------------------------------------------
'
'   ROM書込み進捗表示
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


                '---ﾓｼﾞｭｰﾙ形名書込み確認
'                Call cmm.ReadRom(RTOS_MODULENAME_ADDRESS, 10, s_ReadModuleName())
'                If cmm.CmdState = CMD_ERR Then WriteMode = Error_Write: GoTo WriteEndCheck
'                For i = 0 To 9
'                    If s_ReadModuleName(i) <> WriteModuleNameData(i) Then
'                        WriteMode = Error_Write
'                        GoTo WriteEndCheck
'                    End If
'                Next

