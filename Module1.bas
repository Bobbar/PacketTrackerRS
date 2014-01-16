Attribute VB_Name = "Module1"
Option Explicit
Global cn_global              As New ADODB.Connection
Public Const colCreate        As Long = &HFFBF80 '&H80C0FF
Public Const colInTransit     As Long = &H80FF80
Public Const colReceived      As Long = &HF4FF80 '&H80FFFF
Public Const colClosed        As Long = &H8080FF
Public Const colFiled         As Long = &H8587FF '&HFF8080
Public Const colReopened      As Long = &HFF80FF
Public Const intWaitTime      As Integer = 10000
Public Const strSMTPServer    As String = "mx.wthg.com"
Public Const strServerAddress As String = "10.35.1.40"
Public Const strUsername      As String = "TicketApp"
Public Const strPassword      As String = "yb4w4"
Public strSQLDriver           As String
Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue _
                Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       lpData As Any, _
                                       lpcbData As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (dest As Any, _
                                       Source As Any, _
                                       ByVal numBytes As Long)
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
Const KEY_READ = &H20019 ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
Public Type EmailInfo
    SendOrRec As String
    strFrom As String
    strTo As String
    Comment As String
    JobNum As String
    Description As String
    PartNum As String
    Customer As String
    Creator As String
    CreateDate As String
    GUID As String
    Delivered As Boolean
    TimeStamp As String
End Type
Public Type ReportInfo
    Action As String
    TimeInState As String
    JobNum As String
    Description As String
    Customer As String
    Part As String
    ActionDate As String
    Creator As String
    CreateDate As String
    Color As Long
    Owner As String
End Type
Public ReportData()   As ReportInfo
Public EmailData()    As EmailInfo
Public strUserIndex() As String
Public bolVerbose     As Boolean
Public strLogLoc      As String
Public strCSVLoc      As String
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203 'Double-click
Public Const WM_LBUTTONDOWN = &H201 'Button down
Public Const WM_LBUTTONUP = &H202 'Button up
Public Const WM_RBUTTONDBLCLK = &H206 'Double-click
Public Const WM_RBUTTONDOWN = &H204 'Button down
Public Const WM_RBUTTONUP = &H205 'Button up
Public Declare Function Shell_NotifyIcon _
               Lib "shell32" _
               Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                          pnid As NOTIFYICONDATA) As Boolean
Public Const strDBDateFormat   As String = "YYYY-MM-DD"
Public Const strUserDateFormat As String = "MM/DD/YYYY"
Public strReportHTML           As String
Public dtStartDate             As Date, dtEndDate As Date
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const DayToRun           As Long = vbMonday
Public Const MinutesTillRefresh As Long = 60 'Minutes until user list refresh * 2  (60 = 30)
Public MinsCounted              As Long
Public lngUptime                As Long
Public strAPPTITLE              As String
Public lngAttempts              As Long, lngSuccess As Long, lngRetries As Long
Private Declare Function GetIpAddrTable_API _
                Lib "IpHlpApi" _
                Alias "GetIpAddrTable" (pIPAddrTable As Any, _
                                        pdwSize As Long, _
                                        ByVal bOrder As Long) As Long
Public Function ConnectToDB() As Boolean
    On Error GoTo errs
    ConnectToDB = False
    cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    If cn_global.State = 1 Then
        ConnectToDB = True
    Else
        ConnectToDB = False
    End If
    Exit Function
errs:
    ErrHandle Err.Number, Err.Description, "ConnectToDB"
End Function
' Returns an array with the local IP addresses (as strings).
' Author: Christian d'Heureuse, www.source-code.biz
Public Function GetIpAddrTable()
    Dim Buf(0 To 511) As Byte
    Dim BufSize       As Long: BufSize = UBound(Buf) + 1
    Dim rc            As Long
    rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
    If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
    Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
    If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
    ReDim IpAddrs(0 To NrOfEntries - 1) As String
    Dim i As Integer
    For i = 0 To NrOfEntries - 1
        Dim j As Integer, s As String: s = ""
        For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
        IpAddrs(i) = s
    Next
    GetIpAddrTable = IpAddrs
End Function
Public Function ConvertTime(ByVal lngMS As Long) As String
    Dim lngSeconds As Long, lngDays As Long, lngHours As Long, lngMins As Long
    Dim strSeconds As String, strDays As String
    lngSeconds = lngMS / 1000
    lngDays = Int(lngSeconds / 86400)
    lngSeconds = lngSeconds Mod 86400
    lngHours = Int(lngSeconds / 3600)
    lngSeconds = lngSeconds Mod 3600
    lngMins = Int(lngSeconds / 60)
    lngSeconds = lngSeconds Mod 60
    'If lngSeconds <> 1 Then strSeconds = "s"
    If lngDays <> 1 Then strDays = "s"
    ConvertTime = lngDays & " days,  " & lngHours & ":" & lngMins & ":" & lngSeconds
End Function
Public Sub Wait(ByVal DurationMS As Long)
    Dim EndTime As Long
    EndTime = GetTickCount + DurationMS
    Do While EndTime > GetTickCount
        DoEvents
        Sleep 1
    Loop
End Sub
Public Sub GetUserIndex()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim i       As Integer
    strSQL1 = "select * from users"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset
    i = 1
    ReDim strUserIndex(2, rs.RecordCount)
    Do Until rs.EOF
        With rs
            strUserIndex(0, i) = UCase$(!idUsers)
            strUserIndex(1, i) = !idFullname
            strUserIndex(2, i) = !idEmail
            i = i + 1
            rs.MoveNext
        End With
    Loop
End Sub
Public Function GetEmail(strUsername As String) As String
    Dim i As Integer
    For i = 0 To UBound(strUserIndex, 2)
        If strUserIndex(0, i) = strUsername Then
            GetEmail = UCase$(strUserIndex(2, i))
            Exit Function
        End If
    Next i
End Function
Public Function GetFullName(strUsername As String) As String
    Dim i As Integer
    For i = 0 To UBound(strUserIndex, 2)
        If strUserIndex(0, i) = strUsername Then
            GetFullName = UCase$(strUserIndex(1, i))
            Exit Function
        End If
    Next i
End Function
Public Sub FindMySQLDriver()
    ToLog "Scanning for MySQL Driver..."
    GetODBCDrivers
    Dim i           As Integer
    Dim strPossis() As String
    Dim blah
    ReDim strPossis(0)
    For i = 1 To GetODBCDrivers.Count
        If InStr(1, GetODBCDrivers.Item(i), "MySQL") Then
            strPossis(UBound(strPossis)) = GetODBCDrivers.Item(i)
            ReDim Preserve strPossis(UBound(strPossis) + 1)
        End If
    Next i
    If UBound(strPossis) > 1 Then
        blah = MsgBox("Multiple MySQL Drivers detected!", vbExclamation + vbOKOnly, "Gasp!")
        strSQLDriver = strPossis(0)
    Else
        strSQLDriver = strPossis(0)
    End If
    ToLog "MySQL Driver = " & strSQLDriver
End Sub
Function EnumRegistryValues(ByVal hKey As Long, ByVal KeyName As String) As Collection
    Dim handle            As Long
    Dim Index             As Long
    Dim valueType         As Long
    Dim Name              As String
    Dim nameLen           As Long
    Dim resLong           As Long
    Dim resString         As String
    Dim dataLen           As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal            As Long
    ' initialize the result
    Set EnumRegistryValues = New Collection
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    Do
        ' this is the max length for a key name
        nameLen = 260
        Name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, Index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        ' retrieve the value's name
        valueInfo(0) = Left$(Name, nameLen)
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        Index = Index + 1
    Loop
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
End Function
Function GetODBCDrivers() As Collection
    Dim res    As Collection
    Dim values As Variant
    ' initialize the result
    Set GetODBCDrivers = New Collection
    ' the names of all the ODBC drivers are kept as values
    ' under a registry key
    ' the EnumRegistryValue returns a collection
    For Each values In EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBCINST.INI\ODBC Drivers")
        ' each element is a two-item array:
        ' values(0) is the name, values(1) is the data
        If StrComp(values(1), "Installed", 1) = 0 Then
            ' if installed, add to the result collection
            GetODBCDrivers.Add values(0), values(0)
        End If
    Next
End Function
Public Sub SendNotification(SendRec As String, _
                            MailFrom As String, _
                            MailTo As String, _
                            JobNum As String, _
                            strDescrip As String, _
                            strPartNum As String, _
                            strCustomer As String, _
                            strCreator As String, _
                            strCreateDate As String, _
                            strComment As String, _
                            strTimeStamp As String)
    On Error GoTo errs
    Dim iConf As New CDO.Configuration
    Dim Flds  As ADODB.Fields
    Dim iMsg  As New CDO.Message
    Set Flds = iConf.Fields
    ' Set the configuration
    Flds(cdoSendUsingMethod) = cdoSendUsingPort
    Flds(cdoSMTPServer) = strSMTPServer
    Flds(cdoSMTPServerPort) = 25
    Flds(cdoSMTPConnectionTimeout) = 30
    ' ... other settings
    Flds.Update
    With iMsg
        Set .Configuration = iConf
        .Sender = GetEmail(MailFrom)
        .To = GetEmail(MailTo)
        .From = GetEmail(MailFrom)
        If UCase$(SendRec) = "SEND" Then
            .Subject = "JPT: " & GetFullName(MailFrom) & " sent you a packet  (" & JobNum & ")"
        ElseIf UCase$(SendRec) = "REC" Then
            .Subject = "JPT: " & GetFullName(MailFrom) & " received a packet  (" & JobNum & ")"
        End If
        .HTMLBody = GenerateHTML(SendRec, GetFullName(MailFrom), MailTo, JobNum, strDescrip, strPartNum, strCustomer, strCreator, strCreateDate, strComment, strTimeStamp)
        '.TextBody = Message
        .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Exit Sub
errs:
    ToLog "Failed to send EMail notification!"
    If bolVerbose Then ToLog "ERROR DTL:  SUB = SendNotification | " & Err.Number & " - " & Err.Description
    If Err.Number = -2147220973 Then 'if failed to connect then clear successful deliveries and try again
        ToLog "Could not establish connection. Clearing successful from queue and retrying..."
        lngRetries = lngRetries + 1
        ClearEmailQueue
        CheckQueue
    End If
    If Err.Number = -2147220974 Then 'if lost connection, the email still made it, so fail softly and continue
        ToLog "Lost connection after send, probably... Moving on..."
        lngSuccess = lngSuccess + 1
        Resume Next
    End If
End Sub
Public Function GenerateHTML(strSendOrRec As String, _
                             strFrom As String, _
                             strTo As String, _
                             strPacketNum As String, _
                             strDescrip As String, _
                             strPartNum As String, _
                             strCustomer As String, _
                             strCreator As String, _
                             strCreateDate As String, _
                             strComment As String, _
                             strTimeStamp As String) As String
    On Error GoTo errs
    Dim tmpHTML As String
    If UCase$(strSendOrRec) = "SEND" Then
        Dim BackColor As String
        BackColor = Hex$(colInTransit)
        tmpHTML = tmpHTML + "<HTML>" & vbCrLf
        tmpHTML = tmpHTML + "<BODY BGCOLOR=" & BackColor & ">" & vbCrLf
        tmpHTML = tmpHTML + "<FONT STYLE=font-family:Tahoma;>" & vbCrLf
        tmpHTML = tmpHTML + "<FONT SIZE=2>" & strTimeStamp & "</FONT><BR>"
        tmpHTML = tmpHTML + "<FONT SIZE=6><U>Message from Job Packet Tracker:</U></FONT><BR><BR>" & vbCrLf
        tmpHTML = tmpHTML + strFrom & " is sending Job Packet <b>" & strPacketNum & "</b> to you. <BR><BR>" & vbCrLf
        If strComment <> "" Then
            tmpHTML = tmpHTML + "<I>" & Chr$(34) & strComment & Chr$(34) & "</I><BR><BR><BR><BR>" & vbCrLf
        Else
            tmpHTML = tmpHTML + "<BR><BR>"
        End If
        tmpHTML = tmpHTML + " <FONT STYLE=font-family:Terminal;>" & vbCrLf
        tmpHTML = tmpHTML + " Detailed Info:<BR><BR>" & vbCrLf
        tmpHTML = tmpHTML + "<table border=0 cellpadding=3>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Job Number:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Description:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strPacketNum & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strDescrip & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Part Number:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Customer:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strPartNum & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strCustomer & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Creator:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Create Date:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & GetFullName(strCreator) & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strCreateDate & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "</table>" & vbCrLf
        tmpHTML = tmpHTML + " <FONT>" & vbCrLf
        tmpHTML = tmpHTML + " </BODY>" & vbCrLf
        tmpHTML = tmpHTML + " </HTML>" & vbCrLf
        GenerateHTML = tmpHTML
    ElseIf UCase$(strSendOrRec) = "REC" Then
        BackColor = Hex$(&HF4FF80)
        tmpHTML = tmpHTML + "<HTML>" & vbCrLf
        tmpHTML = tmpHTML + "<BODY BGCOLOR=" & BackColor & ">" & vbCrLf
        tmpHTML = tmpHTML + "<FONT STYLE=font-family:Tahoma;>" & vbCrLf
        tmpHTML = tmpHTML + "<FONT SIZE=2>" & DateTime.Date$ & " " & DateTime.Time$ & "</FONT><BR>" & vbCrLf
        tmpHTML = tmpHTML + "<FONT SIZE=6><U>Message from Job Packet Tracker:</U></FONT><BR><BR>" & vbCrLf
        tmpHTML = tmpHTML + strFrom & " has received Job Packet <b>" & strPacketNum & "</b><BR><BR>" & vbCrLf
        If strComment <> "" Then
            tmpHTML = tmpHTML + "<I>" & Chr$(34) & strComment & Chr$(34) & "</I><BR><BR><BR><BR>" & vbCrLf
        Else
            tmpHTML = tmpHTML + "<BR><BR>"
        End If
        tmpHTML = tmpHTML + " <FONT STYLE=font-family:Terminal;>" & vbCrLf
        tmpHTML = tmpHTML + " Detailed Info:<BR><BR>" & vbCrLf
        tmpHTML = tmpHTML + "<table border=0 cellpadding=3>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Job Number:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Description:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strPacketNum & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strDescrip & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Part Number:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Customer:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strPartNum & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strCustomer & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Creator:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td><b>Create Date:</b></td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & GetFullName(strCreator) & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strCreateDate & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "</table>" & vbCrLf
        tmpHTML = tmpHTML + " <FONT>" & vbCrLf
        tmpHTML = tmpHTML + " </BODY>" & vbCrLf
        tmpHTML = tmpHTML + " </HTML>" & vbCrLf
        GenerateHTML = tmpHTML
    End If
    Exit Function
errs:
    ToLog "Failed to Genterate HTML!"
    If bolVerbose Then ToLog "ERROR DTL:  SUB = GenerateHTML | " & Err.Number & " - " & Err.Description
End Function
Public Sub CheckQueue()
    On Error GoTo errs
    Dim tmpGUID As String
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM emailqueue d LEFT JOIN packetlist c ON c.idJobNum = d.idJobNum"
    Set rs = cn_global.Execute(strSQL1)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
        ToLog rs.RecordCount & " Email(s) found in queue.  Parsing..."
    End If
    ReDim EmailData(rs.RecordCount)
    Do Until rs.EOF
        With rs
            tmpGUID = .Fields(5)
            EmailData(.AbsolutePosition - 1).GUID = tmpGUID
            EmailData(.AbsolutePosition - 1).SendOrRec = !idSendOrRec
            EmailData(.AbsolutePosition - 1).strFrom = !idFrom
            EmailData(.AbsolutePosition - 1).strTo = !idTo
            EmailData(.AbsolutePosition - 1).JobNum = !idJobNum
            EmailData(.AbsolutePosition - 1).PartNum = !idPartNum
            EmailData(.AbsolutePosition - 1).Customer = !idCustPONum
            EmailData(.AbsolutePosition - 1).Comment = !idComment
            EmailData(.AbsolutePosition - 1).Creator = !idCreator
            EmailData(.AbsolutePosition - 1).CreateDate = !idCreateDate
            EmailData(.AbsolutePosition - 1).Description = !idDescription
            EmailData(.AbsolutePosition - 1).Delivered = False
            EmailData(.AbsolutePosition - 1).TimeStamp = !idTimeStamp
            .MoveNext
        End With
    Loop
    rs.Close
    SendEmails
    ClearEmailQueue
    ToLog "Done..."
    Exit Sub
errs:
    ToLog "Error Checking Queue!"
    ErrHandle Err.Number, Err.Description, "CheckQueue"
    If Err.Number = 94 Then
        ToLog "Null values detected! Packet data missing. Clearing single item from queue."
        ClearEmailQueue tmpGUID
    End If
End Sub
Public Sub ErrHandle(lngErrNum As Long, strErrDescription As String, strOrigSub As String)
    Select Case lngErrNum
        Case -2147467259, 3704
            JPTRS.lblStatus.Caption = "Disconnected!"
            If bolVerbose Then ToLog "ERROR DTL:  SUB = CheckQueue | " & Err.Number & " - " & Err.Description
            ToLog "SQL Connection Lost!  Trying to Reconnect..."
            Set cn_global = Nothing
            If ConnectToDB Then
                ToLog "Connected!"
            End If
        Case Else
            ToLog lngErrNum & " - " & strErrDescription & " | " & strOrigSub
    End Select
End Sub
Public Sub ToLog(Message As String)
    Dim tmpMsg As String
    If Dir$(strLogLoc, vbDirectory) = "" Then MkDir strLogLoc 'Environ$("APPDATA") & "\JPTRS\"
    Open strLogLoc & "\LOG.LOG" For Append As #1
    With JPTRS
        tmpMsg = DateTime.Date & " " & DateTime.Time & ": " & Message
        .lstLog.AddItem tmpMsg, 0
        Print #1, tmpMsg
        Close #1
       ' DoEvents
    End With
End Sub
Public Sub SendEmails()
    On Error GoTo errs
    Dim i As Integer
    For i = 0 To UBound(EmailData) - 1
        If bolVerbose Then ToLog "Sending SMTP " & i + 1 & " of " & UBound(EmailData) & " : " & EmailData(i).SendOrRec & " - " & EmailData(i).strFrom & " - " & EmailData(i).strTo & " - " & EmailData(i).JobNum & " - " & EmailData(i).Description & " - " & EmailData(i).PartNum & " - " & EmailData(i).Customer & " - " & EmailData(i).Creator & " - " & EmailData(i).CreateDate & " - " & EmailData(i).Comment
        JPTRS.lblStatus.Caption = "Sending EMail...."
        lngAttempts = lngAttempts + 1
        SendNotification EmailData(i).SendOrRec, EmailData(i).strFrom, EmailData(i).strTo, EmailData(i).JobNum, EmailData(i).Description, EmailData(i).PartNum, EmailData(i).Customer, EmailData(i).Creator, EmailData(i).CreateDate, EmailData(i).Comment, EmailData(i).TimeStamp
        EmailData(i).Delivered = True
        lngSuccess = lngSuccess + 1
        JPTRS.lblStatus.Caption = "Waiting...."
        ToLog "Waiting..."
        Wait intWaitTime 'wait to avoid flooding server
    Next
    Exit Sub
errs:
    ToLog "Error Sending Emails!"
    ToLog "ERROR DTL:  SUB = SendEmails | " & Err.Number & " - " & Err.Description
End Sub
Public Sub ClearEmailQueue(Optional strGUID As String)
  On Error GoTo errs
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    If strGUID = "" Then
        If bolVerbose Then ToLog "Clearing Queue..."
        For i = 0 To UBound(EmailData) - 1
            If EmailData(i).Delivered Then
                strSQL1 = "SELECT * From emailqueue Where idGUID = '" & EmailData(i).GUID & "'"
                JPTRS.lblStatus.Caption = "Clearing Queue..."
                rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
                With rs
                    .Delete
                    .Update
                End With
                rs.Close
            End If
        Next i
        ReDim EmailData(0)
    Else
        If bolVerbose Then ToLog "Clearing GUID " & strGUID & " from Queue..."
        strSQL1 = "SELECT * From emailqueue Where idGUID = '" & strGUID & "'"
        rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
        With rs
            .Delete
            .Update
        End With
        rs.Close
    End If
    Exit Sub
errs:
    ToLog "Error Clearing Queue!"
    ToLog "ERROR DTL:  SUB = ClearEmailQueue | " & Err.Number & " - " & Err.Description
End Sub
Public Sub WeeklyReportGetData()
    Dim lngTIS  As Long
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    i = 0
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM packetlist d LEFT JOIN packetentrydb c ON c.idJobNum = d.idJobNum WHERE idDate = (SELECT MAX(idDate) FROM packetentrydb c2 Where c2.idJobNum = d.idJobNum) AND idStatus='OPEN' ORDER BY idDate DESC LIMIT 50"
    ToLog "Starting Weekly Report..."
    Set rs = cn_global.Execute(strSQL1)
    dtStartDate = DateAdd("d", -7, DateTime.Date)
    dtEndDate = DateAdd("d", -3, DateTime.Date)
    If rs.RecordCount < 1 Then Exit Sub
    ReDim ReportData(0)
    With rs
        Do Until .EOF
            If Format(!idDate, strUserDateFormat) >= dtStartDate And Format(!idDate, strUserDateFormat) <= dtEndDate Then
                If !idAction = "CREATED" Then
                    ReportData(i).Color = colCreate
                ElseIf !idAction = "INTRANSIT" Then
                    ReportData(i).Color = colInTransit
                ElseIf !idAction = "RECEIVED" Then
                    ReportData(i).Color = colReceived
                ElseIf !idAction = "CLOSED" Then
                    ReportData(i).Color = colClosed
                ElseIf !idStatus = "OPEN" And !idAction = "FILED" Then
                    ReportData(i).Color = colFiled
                ElseIf !idAction = "REOPENED" Then
                    ReportData(i).Color = colReopened
                End If
                If !idAction = "CREATED" Then
                    ReportData(i).Action = !idAction & " by " & GetFullName(!idCreator)
                ElseIf !idAction = "INTRANSIT" Then
                    ReportData(i).Action = !idAction & " to " & GetFullName(!idUserTo)
                ElseIf !idAction = "RECEIVED" Then
                    ReportData(i).Action = !idAction & " by " & GetFullName(!idUserFrom)
                ElseIf !idAction = "NULL" Then
                    ReportData(i).Action = "CLOSED by " & GetFullName(!idUser)
                ElseIf !idStatus = "OPEN" And !idAction = "FILED" Then
                    ReportData(i).Action = !idAction & " by " & GetFullName(!idUser)
                ElseIf !idAction = "REOPENED" Then
                    ReportData(i).Action = !idAction & " by " & GetFullName(!idUser)
                End If
                lngTIS = DateDiff("n", !idDate, Date & " " & Time)
                ReportData(i).TimeInState = (IIf(lngTIS > 1440, Round(lngTIS / 1440, 1) & " days ", Round(lngTIS / 60, 1) & " hrs "))
                ReportData(i).ActionDate = !idDate
                ReportData(i).CreateDate = !idCreateDate
                ReportData(i).Creator = !idCreator
                ReportData(i).Customer = !idCustPONum
                ReportData(i).Description = !idDescription
                ReportData(i).JobNum = !idJobNum
                ReportData(i).Part = !idPartNum
                i = i + 1
                ReDim Preserve ReportData(UBound(ReportData) + 1)
            End If
            .MoveNext
        Loop
    End With
    'WeeklyReportParseCSV
    WeeklyReportParseHTML
    SendReport "JPTReportServer@worthingtonindustries.com", ReportRecpts
End Sub
Public Sub WeeklyReportParseHTML()
    On Error GoTo errs
    ToLog "Parsing report to HTML..."
    Dim tmpHTML As String
    Dim i       As Integer
    tmpHTML = tmpHTML + "<FONT STYLE=font-family:Tahoma;>" & vbCrLf
    tmpHTML = tmpHTML + "<FONT SIZE=4><U>Weekly Job Packet Report</U></FONT><BR>" & vbCrLf
    tmpHTML = tmpHTML + "<FONT SIZE=2> Run Date: " & Now & "</FONT><BR>" & vbCrLf
    tmpHTML = tmpHTML + "<FONT SIZE=2> Report Date: " & dtStartDate & " to " & dtEndDate & "</FONT><BR><BR>" & vbCrLf
    tmpHTML = tmpHTML + "<table border=1 cellpadding=2>" & vbCrLf
    tmpHTML = tmpHTML + "<tr>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Job Num</th>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Action</th>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Action Date</th>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Time In State</th>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Description</th>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Customer</th>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Create Date</th>" & vbCrLf
    tmpHTML = tmpHTML + "<th bgcolor=B3B3B3>Creator</th>" & vbCrLf
    tmpHTML = tmpHTML + "</tr>"
    For i = 0 To UBound(ReportData) - 1
        tmpHTML = tmpHTML + "<tr>" & vbCrLf
        tmpHTML = tmpHTML + "<td bgcolor=" & Hex$(ReportData(i).Color) & "><B>" & ReportData(i).JobNum & "</B></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td bgcolor=" & Hex$(ReportData(i).Color) & "> <font size=2>" & ReportData(i).Action & "</font></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td bgcolor=" & Hex$(ReportData(i).Color) & "><font size=2>" & ReportData(i).ActionDate & "</font></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td align=center bgcolor=" & Hex$(ReportData(i).Color) & ">" & ReportData(i).TimeInState & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td bgcolor=" & Hex$(ReportData(i).Color) & "><font size=2>" & ReportData(i).Description & "</font></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td bgcolor=" & Hex$(ReportData(i).Color) & ">" & ReportData(i).Customer & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td bgcolor=" & Hex$(ReportData(i).Color) & "><font size=2>" & ReportData(i).CreateDate & "</font></td>" & vbCrLf
        tmpHTML = tmpHTML + "<td bgcolor=" & Hex$(ReportData(i).Color) & ">" & GetFullName(ReportData(i).Creator) & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>"
    Next
    tmpHTML = tmpHTML + "</table>"
    strReportHTML = tmpHTML
    Exit Sub
errs:
    ToLog "Failed while parsing Weekly Report HTML!"
    If bolVerbose Then ToLog "ERROR DTL:  SUB =  WeeklyReportParseHTML | " & Err.Number & " - " & Err.Description
End Sub
Public Sub SendReport(MailFrom As String, MailTo As String)
    ToLog "Sending Weekly Report...  (" & dtStartDate & " to " & dtEndDate & ")"
    On Error GoTo errs
    Dim iConf As New CDO.Configuration
    Dim Flds  As ADODB.Fields
    Dim iMsg  As New CDO.Message
    Set Flds = iConf.Fields
    ' Set the configuration
    Flds(cdoSendUsingMethod) = cdoSendUsingPort
    Flds(cdoSMTPServer) = strSMTPServer
    Flds(cdoSMTPConnectionTimeout) = 30
    ' ... other settings
    Flds.Update
    With iMsg
        Set .Configuration = iConf
        .Sender = MailFrom 'GetEmail(MailFrom)
        .To = MailTo 'GetEmail(MailTo)
        .From = MailFrom 'GetEmail(MailFrom)
        .Subject = "JPT: Weekly Report (" & dtStartDate & " to " & dtEndDate & ")"
        .HTMLBody = strReportHTML
        '.TextBody = Message
        .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    'Set registry flag telling me that the report for this week has been sent
    SaveSetting App.EXEName, "WeeklyReportSent", "Sent", "TRUE"
    ToLog "Report Sent..."
    Exit Sub
errs:
    ToLog "Failed to send Weekly Report!"
    If bolVerbose Then ToLog "ERROR DTL:  SUB = SendReport | " & Err.Number & " - " & Err.Description
End Sub
Public Function ReportRecpts() As String
    ReportRecpts = ""
    Dim tmpRecpts As String
    Dim rs        As New ADODB.Recordset
    Dim strSQL1   As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM users WHERE idJPTReport = '1'"
    Set rs = cn_global.Execute(strSQL1)
    With rs
        Do Until .EOF
            tmpRecpts = tmpRecpts + !idEmail + ";"
            .MoveNext
        Loop
    End With
    ReportRecpts = tmpRecpts
End Function
Public Function OKToRun() As Boolean
    Dim Flag As Boolean
    Flag = CBool(GetSetting(App.EXEName, "WeeklyReportSent", "Sent", 0))
    OKToRun = False
    If Weekday(Now) = DayToRun And Flag Then
        OKToRun = False
    ElseIf Weekday(Now) <> DayToRun And Flag Then
        OKToRun = False
        SaveSetting App.EXEName, "WeeklyReportSent", "Sent", "FALSE"
    ElseIf Weekday(Now) <> DayToRun And Not Flag Then
        OKToRun = False
    ElseIf Weekday(Now) = DayToRun And Not Flag Then
        OKToRun = True
    End If
    JPTRS.lblReportStatus.Caption = Str(Flag)
End Function
Public Function strDayOfWeek(intDayOfWeek As Integer) As String
    Select Case intDayOfWeek
        Case 1
            strDayOfWeek = "Sunday"
        Case 2
            strDayOfWeek = "Monday"
        Case 3
            strDayOfWeek = "Tuesday"
        Case 4
            strDayOfWeek = "Wednesday"
        Case 5
            strDayOfWeek = "Thursday"
        Case 6
            strDayOfWeek = "Friday"
        Case 7
            strDayOfWeek = "Saturday"
    End Select
End Function
Public Sub WeeklyReportParseCSV()
    On Error GoTo errs
    Dim strCSVName As String, strCSVFullName As String
    Dim i          As Integer
    strCSVName = Format$(DateTime.Date & " " & DateTime.Time, "YYYY-MM-DD-hh-mm-ss") & ".csv"
    strCSVFullName = strCSVLoc & strCSVName
    Open strCSVFullName For Append As #2
    For i = 0 To UBound(ReportData)
        Print #2, ReportData(i).Action & "," & ReportData(i).JobNum & "," & ReportData(i).Description & "," & ReportData(i).Customer & "," & ReportData(i).JobNum & "," & ReportData(i).Creator & "," & ReportData(i).CreateDate
    Next
    Close #2
    Exit Sub
errs:
    ToLog "Failed while parsing Weekly Report!"
    If bolVerbose Then ToLog "ERROR DTL:  SUB = WeeklyReportParseCSV | " & Err.Number & " - " & Err.Description
End Sub
