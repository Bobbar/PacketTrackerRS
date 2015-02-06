Attribute VB_Name = "SocketModule"
Public strSocketData       As String
Public strSocketRequestID  As String
Public strSocketAcceptedID As String
Private Type PacketType
    ID As String
    Type As String
    DataString As String
End Type
Private PacketData           As PacketType
Public Const CommandPacket   As String = "COM"
Public Const RequestPacket   As String = "REQ"
Public Const TerminatePacket As String = "TERM"
Public Const PasswordPacket  As String = "PWD"
Public Const LogPacket       As String = "LOG"
Public Const NamePacket      As String = "NAME"
Public bolWaitingForPass     As Boolean
Public Sub ParsePacket(Data As String)
    Dim SplitData
    SplitData = Split(Data, ",")
    PacketData.ID = SplitData(0)
    PacketData.Type = SplitData(1)
    PacketData.DataString = SplitData(2)
    HandlePacket PacketData
End Sub
Public Function AuthPacket(Packet As PacketType) As Boolean
    AuthPacket = False
    If Packet.ID = strSocketAcceptedID Then AuthPacket = True
End Function
Public Sub HandlePacket(Packet As PacketType)
    Select Case Packet.Type
        Case CommandPacket
            If AuthPacket(Packet) Then PacketCommand Packet.DataString
        Case RequestPacket
        Case TerminatePacket
        Case PasswordPacket
            If bolWaitingForPass Then CheckPassword Packet.DataString
        Case NamePacket
            Logger "TCP Socket: Computer name = " & Packet.DataString
            strSocketRequestID = Packet.DataString
            strSocketAcceptedID = Packet.DataString
    End Select
End Sub
Public Sub PacketCommand(Command As String)
    Logger "Remote Command From " & strSocketAcceptedID & ": " & Command
    Command = UCase$(Command)
    Select Case Command
        Case "UPDATEUSERLIST"
            Logger "Updating user list..."
            RefreshUserList
        Case "CLEARQUEUE"
            ClearEmailQueueAll
        Case "UPTIME"
            Logger ConvertTime(DateTime.Now)
        Case "STARTREPORT DAILY"
        Case "STARTREPORT WEEKLY"
        Case "PAUSE"
            Logger "Pausing exeution..."
            With JPTRS
                .tmrCheckQueue.Enabled = False
                .tmrReportClock.Enabled = False
                .tmrTaskTimer.Enabled = False
                bolExecutionPaused = True
            End With
        Case "RESUME"
            Logger "Resuming exeution..."
            With JPTRS
                .tmrCheckQueue.Enabled = True
                .tmrReportClock.Enabled = True
                .tmrTaskTimer.Enabled = True
                bolExecutionPaused = False
            End With
        Case "ENDPROGRAM"
            Logger "Ending program..."
            Wait 1000
            EndProgram
        Case "STATUS"
            Logger "STATUS: Uptime: " & ConvertTime(DateTime.Now) & "    Atmpts, Sucss, Rtry: " & lngAttempts & ", " & lngSuccess & ", " & lngRetries
        Case "PASSWORD"
            CheckPassword Command
        Case "LOADLOG"
            SendLog
        Case Else
            Logger "'" & Command & "' is not a recognized command."
    End Select
End Sub
Public Sub SendLog()
    Dim i As Integer
    With JPTRS
        SocketLog "[LOG START]"
        For i = .lstLog.ListCount - 1 To 0 Step -1
            SocketLog .lstLog.List(i)
        Next i
        SocketLog "[LOG END]"
    End With
End Sub
Public Sub CheckPassword(Password As String)
    On Error GoTo errs
    Dim rs          As New ADODB.Recordset
    Dim strSQL1     As String
    Dim strPassword As String
    strSQL1 = "SELECT * FROM socketvars"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset
    With rs
        strPassword = !idPassword
    End With
    If Password = strPassword Then
        AcceptPassword
        Logger "TCP Socket: Password accepted!"
        strSocketAcceptedID = PacketData.ID
    Else
        RejectPassword
        Logger "TCP Socket: Password rejected!"
        strSocketAcceptedID = vbNullString
    End If
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "CheckPassword"
End Sub
Public Sub AcceptPassword()
    SendData strSocketRequestID & "," & PasswordPacket & ",GOODPASS"
    bolWaitingForPass = False
End Sub
Public Sub RejectPassword()
    SendData strSocketRequestID & "," & CommandPacket & ",BADPASS"
    bolWaitingForPass = False
End Sub
Public Sub RequestPass()
    SocketLog "Password?"
    SendData strSocketRequestID & "," & PasswordPacket & ",GIVEPASS"
    bolWaitingForPass = True
End Sub
Public Sub RequestName()
    SocketLog "Computer name?"
    SendData strSocketRequestID & "," & RequestPacket & ",GIVENAME"
    bolWaitingForPass = True
End Sub
Public Sub SendData(Data As String)
    With JPTRS
        If .TCPServer.State <> sckClosed Then
            .TCPServer.SendData Chr$(1) & Data
        End If
    End With
End Sub
Public Sub SocketLog(strLog As String)
    SendData strSocketRequestID & "," & LogPacket & "," & strLog
End Sub
Public Sub StartTCPServer()
    On Error GoTo errs
    JPTRS.TCPServer.LocalPort = strListenPort
    JPTRS.TCPServer.Listen
    Logger "Listening on port " & strListenPort
    Exit Sub
errs:
    Logger "***** Error Starting TCP Server! *****"
    ErrHandle Err.Number, Err.Description, "StartTCPServer"
End Sub
