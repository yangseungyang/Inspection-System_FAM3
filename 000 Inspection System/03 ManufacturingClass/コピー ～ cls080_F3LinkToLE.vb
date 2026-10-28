Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : cls080_F3LinkToLE
'Function   : Communication F3LE etc 
'Language   : VB2008　Framework3.5
'Reference  : 
'Memo       :
'           :
'Issue      : 2009/02/  Rev.0    T.Yanagawa       New Release
'============================================================================================================================
'Doevents
Imports System.Windows.Forms.Application

Public Class cls080_F3LinkToLE

    'Local ================================================
    Dim mPort As String
    Dim mSetting As String
    Dim mTerminalChar As String
    Dim mRecvData As String
    Dim mErrorComment As String
    Dim mErrorCode As String

    Dim mIPAdd As String

    Private intPortNo As Integer = 12289
    Private objNetStream As System.Net.Sockets.NetworkStream
    Private objTcpClient As System.Net.Sockets.TcpClient

    'Property =====================================================================================
    'IPAddress -----------------------------
    Public WriteOnly Property IPAdd() As String
        Set(ByVal Value As String)
            mIPAdd = Value
        End Set
    End Property

    'Receive ---------------------------
    Public ReadOnly Property RECEIVE_DATA() As String
        Get
            RECEIVE_DATA = mRecvData
        End Get
    End Property

    'Error ----------------------------
    Public ReadOnly Property ERR_COMMENT() As String
        Get
            ERR_COMMENT = mErrorComment
        End Get
    End Property

    Public ReadOnly Property ERR_CODE() As String
        Get
            ERR_CODE = mErrorCode
        End Get
    End Property

    Public Sub F3PCLinkCommunication(ByVal SendCommand As String, ByRef RecvCommand As String)
        '==========================================================================================
        ' F.Name   : F3PCLinkCommunication
        ' Function : PC Link Command Communication
        ' Data     : Send command
        ' Return   : Recv command
        ' Remarks  :
        '==========================================================================================

        Dim bytReceiveBuffer(255) As Byte
        Dim bytSendBuffer As Byte()
        Dim intDataLength As Integer
        Dim strReceivedData As String = ""

        objTcpClient = New System.Net.Sockets.TcpClient
        With objTcpClient

            Try

                objTcpClient.Connect(mIPAdd, intPortNo)
                objNetStream = .GetStream
                With objNetStream
                    bytSendBuffer = System.Text.Encoding.Default.GetBytes(SendCommand & vbCrLf)
                    .Write(bytSendBuffer, 0, bytSendBuffer.Length)

                    mWait(0.1)

                    Do
                        If .DataAvailable Then
                            intDataLength = .Read(bytReceiveBuffer, 0, bytReceiveBuffer.Length)
                            strReceivedData &= System.Text.Encoding.Default.GetString(bytReceiveBuffer).Substring(0, intDataLength)
                        ElseIf Not strReceivedData Is Nothing Then
                            RecvCommand = strReceivedData
                            mRecvData = strReceivedData
                            Exit Do
                        End If
                    Loop

                End With

            Catch Ex As Exception
                mErrorComment = "PC Link Communication Error"
                Throw New ApplicationException(mErrorComment)

            Finally
                .Close()

            End Try

        End With

        objTcpClient = Nothing

    End Sub

    Private Sub mWait(ByVal Sec As Double)
        '==========================================================================================
        ' F.Name   : mWait
        ' Function : Sec Timer
        ' Data     : 0.000 sec
        ' Return   : Nothing
        ' Remarks  :
        '==========================================================================================

        Dim StartTime As Double
        Dim NowTime As Double
        Dim ErrTime As Double

        StartTime = Second(Now)
        Do While Sec > ErrTime
            NowTime = Second(Now)
            ErrTime = NowTime - StartTime
            If ErrTime < 0 Then ErrTime = 24.0# * 3600.0# - StartTime + NowTime
            DoEvents()
        Loop

    End Sub


End Class
