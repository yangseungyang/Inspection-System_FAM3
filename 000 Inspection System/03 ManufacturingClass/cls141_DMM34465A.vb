Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : cls141_DMM34465A
'Function   : 전압 측정 - A8 설비의 DMM 34465A(GPIB) 로 측정 (cls140_VoltageMeasure 의 DMM 7555 측정 대체)
'Language   : VB2019　Framework4.8
'Memo       : ----------------------------- Ver0.52 YSH ADD (클래스 전체 신규) -----------------------------
'             - A8 프로그램의 계측기 클래스를 수정 없이 그대로 사용한다
'                 03 ManufacturingClass\cls103_CommMain.vb (GPIB 연결 : Setup ini 의 <DEVICEid>=0 → NI GPIB, <PARAMid>=GPIB 주소)
'                 03 ManufacturingClass\cls240_34410A.vb   (34410A / 34465A 제어)
'                 010 EquipmentData\Drive\Setup34410A.ini  (A8 과 같은 파일)
'             - 측정 순서는 A8 cls2X0_A8HCA_inspection.DMMRead 와 같다 (New → Open → 레인지 설정 → GetMeasureDMM → Close)
'             - 레인지 : K호기 DMM 7555 는 DCV 20V 고정 → 34465A 는 DCV 10V 레인지 (측정 대상 최대 VCC 5V)
'Issue      : 2026/10/27  Ver0.52    YSH       New Release
'             근거 : FAM3 증산 대응_CPU검사장치 ADD-ON.md Rev 2  7-1, 7-2 / CPU검사장치_Router대응_분석_v1.1.xlsx 6_대응계획 No 7
'================================================================================================================================

Public Class cls141_DMM34465A

    Private Const DMM_ID As Integer = 1             'Setup34410A.ini 의 <DEVICE1> / <PARAM1> (A8 DMMRead 와 같은 번호)
    Private Const WAIT_SEC As Double = 0            'GetMeasureDMM 대기 (A8 DMMRead 와 같음)

    Public Shared Function MeasureDCV() As Single
        '==========================================================================================
        'Function : DC 전압 1회 측정
        'Return   : 측정값 (V)
        '==========================================================================================
        Dim cDMM As cls240_34410A = Nothing

        Try
            cDMM = New cls240_34410A(DMM_ID)
            cDMM.Open()
            cDMM.SetDCVoltageRange(cls240_34410A.DCVoltageRange.V10V)
            System.Threading.Thread.Sleep(200)
            Return CSng(cDMM.GetMeasureDMM(WAIT_SEC))

        Catch ex As Exception
            Throw New ApplicationException("DMM34465A測定異常 : " & ex.Message)

        Finally
            If cDMM IsNot Nothing Then
                Try
                    cDMM.Close()
                Catch
                End Try
            End If

        End Try

    End Function

End Class
