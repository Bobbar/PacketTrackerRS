Attribute VB_Name = "Module1"
Option Explicit
Global cn_global              As New ADODB.Connection
Public Const intWaitTime      As Integer = 10000
Public Const strSMTPServer    As String = "mx.wthg.com"
Public Const strServerAddress As String = "ohbre-pwadmin01"
Public Const strUsername      As String = "TicketApp"
Public Const strPassword      As String = "yb4w4"
Public Const strListenPort    As String = "1001"
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
Public bolVerbose As Boolean
Public strLogLoc  As String
Public strCSVLoc  As String
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
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Const MinutesTillRefresh      As Long = 5 'Minutes between user list refresh
Public Const MinutesTillStatusReport As Long = 720 'Minutes between status updates in log
Public MinsCounted                   As Long
Public strAPPTITLE                   As String
Public lngAttempts                   As Long, lngSuccess As Long, lngRetries As Long
Public lngStartTime                  As Long
Public lngCurTime                    As Long
Private Declare Function GetIpAddrTable_API _
                Lib "IpHlpApi" _
                Alias "GetIpAddrTable" (pIPAddrTable As Any, _
                                        pdwSize As Long, _
                                        ByVal bOrder As Long) As Long
Public intRetryFail          As Integer
Public Const intRetryFailMax As Integer = 5
Public bolExecutionPaused    As Boolean
Public Type UserAttributes
    UserName As String
    FullName As String
    EMail As String
    GetsDaily As Boolean
    GetsWeekly As Boolean
    Filters As String
End Type
Public Users()      As UserAttributes
Public strLogBuffer As String
Public Function StatusReport() As String
StatusReport = "STATUS: Uptime: " & ConvertTime(DateTime.Now) & "   Version: " & App.Major & "." & App.Minor & "." & App.Revision & "    Atmpts, Sucss, Rtry: " & lngAttempts & ", " & lngSuccess & ", " & lngRetries
End Function
Public Sub RefreshUserList()
    With JPTRS
        .tmrCheckQueue.Enabled = False
        .tmrReportClock.Enabled = False
        GetUserIndex
        .tmrCheckQueue.Enabled = True
        .tmrReportClock.Enabled = True
    End With
End Sub
Public Sub CheckQueue()
    On Error GoTo errs
    Dim tmpGUID As String
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim intPos  As Integer
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM emailqueue d LEFT JOIN packetlist c ON c.idJobNum = d.idJobNum"
    Set rs = cn_global.Execute(strSQL1)
    If rs.RecordCount = 0 Then
        Exit Sub
    Else
        Logger rs.RecordCount & " Email(s) found in queue.  Parsing..."
    End If
    ReDim EmailData(rs.RecordCount)
    Do Until rs.EOF
        With rs
            tmpGUID = .Fields(5)
            intPos = .AbsolutePosition - 1
            EmailData(intPos).GUID = tmpGUID
            EmailData(intPos).SendOrRec = !idSendOrRec
            EmailData(intPos).strFrom = !idFrom
            EmailData(intPos).strTo = !idTo
            EmailData(intPos).JobNum = !idJobNum
            EmailData(intPos).PartNum = !idPartNum
            EmailData(intPos).Customer = !idCustPONum
            EmailData(intPos).Comment = !idComment
            EmailData(intPos).Creator = !idCreator
            EmailData(intPos).CreateDate = !idCreateDate
            EmailData(intPos).Description = !idDescription
            EmailData(intPos).Status.Sent = False
            EmailData(intPos).Status.Trash = False
            EmailData(intPos).TimeStamp = !idTimeStamp
            .MoveNext
        End With
    Loop
    rs.Close
    SendEmails
    ClearEmailQueue
    Logger "Done..."
    Exit Sub
errs:
    Logger "***** Error Checking Queue! *****"
    ErrHandle Err.Number, Err.Description, "CheckQueue"
    If Err.Number = 94 Then
        Logger "***** Null or invalid values detected! Packet data corrupt. Clearing single item from queue. *****"
        ClearEmailQueue tmpGUID
        CheckQueue
    End If
End Sub
Public Sub FixEmailQueue()
    On Error GoTo errs
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    If bolVerbose Then Logger "Fixing Queue..."
    For i = 0 To UBound(EmailData) - 1
        If Not EmailData(i).Status.Sent Then
            If EmailData(i).JobNum = "" Or EmailData(i).SendOrRec = "" Or EmailData(i).strFrom = "" Or EmailData(i).strTo = "" Then
                strSQL1 = "SELECT * From emailqueue Where idGUID = '" & EmailData(i).GUID & "'"
                Logger "Clearing corrupt entry..." & "  GUID: " & EmailData(i).GUID
                rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
                rs.Delete
                rs.Update
                rs.Close
            End If
        End If
    Next i
    ReDim EmailData(0)
    Exit Sub
errs:
    Logger "Error Fixing Queue!"
    Logger "ERROR DTL:  SUB = FixEmailQueue | " & Err.Number & " - " & Err.Description
End Sub
Public Sub ClearEmailQueue(Optional strGUID As String)
    On Error GoTo errs
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    If strGUID = "" Then
        If bolVerbose Then Logger "Clearing Queue..."
        For i = 0 To UBound(EmailData) - 1
            If EmailData(i).Status.Sent Or EmailData(i).Status.Trash Then
                strSQL1 = "SELECT * From emailqueue Where idGUID = '" & EmailData(i).GUID & "'"
                JPTRS.lblStatus.Caption = "Clearing Queue..."
                rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
                rs.Delete
                rs.Update
                rs.Close
            End If
        Next i
        ReDim EmailData(0)
    Else
        If bolVerbose Then Logger "Clearing GUID " & strGUID & " from Queue..."
        strSQL1 = "SELECT * From emailqueue Where idGUID = '" & strGUID & "'"
        rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
        With rs
            .Delete
            .Update
        End With
        rs.Close
        ReDim EmailData(0)
    End If
    Exit Sub
errs:
    Logger "Error Clearing Queue!"
    Logger "ERROR DTL:  SUB = ClearEmailQueue | " & Err.Number & " - " & Err.Description
End Sub
Public Sub ClearEmailQueueAll()
    On Error GoTo errs
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    If bolVerbose Then Logger "Force Clearing Queue..."
    'For i = 0 To UBound(EmailData) - 1
    strSQL1 = "SELECT * From emailqueue"
    JPTRS.lblStatus.Caption = "Clearing Queue..."
    rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
    Do Until rs.EOF
        rs.Delete
        rs.MoveNext
    Loop
    rs.Update
    rs.Close
    'Next i
    ReDim EmailData(0)
    Exit Sub
errs:
    Logger "Error Clearing Queue!"
    Logger "ERROR DTL:  SUB = ClearEmailQueue | " & Err.Number & " - " & Err.Description
End Sub
Public Sub EndProgram()
    On Error GoTo errs
    Dim blah
    Logger "Closing Server Application..."
    Logger "Closing Connections..."
    JPTRS.TCPServer.Close
    Logger "TCP Socket Closed..."
    cn_global.Close
    Logger "Global ADO Connection Closed..."
    Logger "Unloaded Form..."
    Logger "Goodbye..."
    End
    Exit Sub
errs:
    Logger Err.Number & " - " & Err.Description
    Resume Next
End Sub
Public Function RemoveFromArray(strGUID As String)
    Dim i          As Integer
    Dim tmpArray() As EmailInfo
    For i = 0 To UBound(EmailData)
        If EmailData(i).GUID <> strGUID Then
            ReDim Preserve tmpArray(UBound(tmpArray) + 1)
            tmpArray(UBound(tmpArray)) = EmailData(i)
        End If
    Next i
    ReDim EmailData(UBound(tmpArray))
    EmailData = tmpArray
End Function
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
Public Function ConvertTime(ByVal CurDate As Date) As String
    Dim lngSeconds As Long, lngDays As Long, lngHours As Long, lngMins As Long
    Dim strSeconds As String, strDays As String
    lngSeconds = DateDiff("s", dtRunDate, CurDate)
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
Public Sub ErrHandle(lngErrNum As Long, strErrDescription As String, strOrigSub As String)
    Select Case lngErrNum
        Case -2147467259, 3704
            JPTRS.lblStatus.Caption = "Disconnected!"
            If bolVerbose Then Logger "ERROR DTL:  SUB = CheckQueue | " & Err.Number & " - " & Err.Description
            Wait 5000
            Logger "SQL Connection Lost!  Trying to Reconnect..."
            Set cn_global = Nothing
            If ConnectToDB Then
                Logger "Connected!"
            End If
        Case 94
            Logger lngErrNum & " - " & strErrDescription & " | " & strOrigSub
        Case Else
            Logger "######### Unhandled error! ###########"
            Logger lngErrNum & " - " & strErrDescription & " | " & strOrigSub
            Logger "Ending..."
            Call JPTRS.Form_QueryUnload(0, 0)
    End Select
End Sub
Public Sub FindMySQLDriver()
    Logger "Scanning for MySQL Driver..."
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
    Logger "MySQL Driver = " & strSQLDriver
End Sub
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
Public Sub GetUserIndex()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim i       As Integer
    strSQL1 = "SELECT * FROM users"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset
    i = 1
    ReDim Users(rs.RecordCount)
    Do Until rs.EOF
        With rs
            Users(.AbsolutePosition).UserName = UCase$(!idUsers)
            Users(.AbsolutePosition).FullName = !idFullname
            Users(.AbsolutePosition).EMail = !idEmail
            Users(.AbsolutePosition).GetsDaily = CBool(!idJPTDailyReport)
            Users(.AbsolutePosition).GetsWeekly = CBool(!idJPTReport)
            Users(.AbsolutePosition).Filters = !idCompanyFilters
            rs.MoveNext
        End With
    Loop
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "GetUserIndex"
End Sub
Public Sub SendEmails()
    On Error GoTo errs
    Dim tmpEmailData As EmailInfo 'temp array for procedure calls, to prevent locking the EmailData array
    Dim i            As Integer
    Dim bolDelivered As Boolean
    For i = 0 To UBound(EmailData) - 1
        If bolVerbose Then Logger "Sending SMTP " & i + 1 & " of " & UBound(EmailData) & " : " & EmailData(i).TimeStamp & " - " & EmailData(i).SendOrRec & " - " & EmailData(i).strFrom & " - " & EmailData(i).strTo & " - " & EmailData(i).JobNum & " - " & EmailData(i).Description & " - " & EmailData(i).PartNum & " - " & EmailData(i).Customer & " - " & EmailData(i).Creator & " - " & EmailData(i).CreateDate & " - " & EmailData(i).Comment
        JPTRS.lblStatus.Caption = "Sending EMail...."
        lngAttempts = lngAttempts + 1
        tmpEmailData = EmailData(i)
        Do
            EmailData(i).Status = SendNotification(tmpEmailData.SendOrRec, tmpEmailData.strFrom, tmpEmailData.strTo, tmpEmailData.JobNum, tmpEmailData.Description, tmpEmailData.PartNum, tmpEmailData.Customer, tmpEmailData.Creator, tmpEmailData.CreateDate, tmpEmailData.Comment, tmpEmailData.TimeStamp, tmpEmailData.GUID)
            If Not EmailData(i).Status.Sent And Not EmailData(i).Status.Trash Then Exit For
        Loop Until EmailData(i).Status.Sent Or EmailData(i).Status.Trash
        If EmailData(i).Status.Sent Then
            Logger "Email Sent!"
            lngSuccess = lngSuccess + 1
        End If
        Wait intWaitTime 'wait to avoid flooding server
    Next
    Exit Sub
errs:
    Logger "***** Error Sending Emails! *****"
    Logger "***** ERROR DTL:  SUB = SendEmails | " & Err.Number & " - " & Err.Description
End Sub
Public Function SendNotification(SendRec As String, _
                                 MailFrom As String, _
                                 MailTo As String, _
                                 JobNum As String, _
                                 strDescrip As String, _
                                 strPartNum As String, _
                                 strCustomer As String, _
                                 strCreator As String, _
                                 strCreateDate As String, _
                                 strComment As String, _
                                 strTimeStamp As String, _
                                 strGUID As String, _
                                 Optional bolRetry As Boolean) As NotificationReturnType
    On Error GoTo errs
    Dim iConf As New CDO.Configuration
    Dim Flds  As ADODB.Fields
    Dim iMsg  As New CDO.Message
    SendNotification.Sent = False
    SendNotification.Trash = False
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
    SendNotification.Sent = True
    SendNotification.Trash = True
    Exit Function
errs:
    Logger "***** Failed to send EMail notification! ******"
    If bolVerbose Then Logger "*****ERROR DTL:  SUB = SendNotification | " & Err.Number & " - " & Err.Description
    If Err.Number = -2147220980 Or Err.Number = -2147220979 Then
        Logger "Email address(s) not found. Waiting, Refreshing user index and trying again..."
        Wait 5000
        GetUserIndex
        Logger "User Index Refreshed..."
        Logger "Retrying... " & intRetryFail + 1 & " of " & intRetryFailMax
        intRetryFail = intRetryFail + 1
        If intRetryFail >= intRetryFailMax Then
            intRetryFail = 0
            Logger "That's enough of that. Marking packet for deletion and moving on."
            Logger SendRec & " - " & MailFrom & " - " & MailTo & " - " & JobNum & " - " & strDescrip & " - " & strPartNum & " - " & strCustomer & " - " & strCreator & " - " & strCreateDate & " - " & strComment & " - " & strTimeStamp
            SendNotification.Sent = False
            SendNotification.Trash = True
            Exit Function
        End If
    End If
    If Err.Number = -2147220973 Then 'if failed to connect then clear successful deliveries and try again
        Logger "***** Could not establish connection. Restarting... *****"
        SendNotification.Sent = False
        SendNotification.Trash = False
        lngRetries = lngRetries + 1
        Wait intWaitTime
        Exit Function
    End If
    If Err.Number = -2147220974 Then 'if lost connection, the email still made it, so fail softly and continue
        Logger "***** Lost connection after send, probably... Moving on... *****"
        lngSuccess = lngSuccess + 1
        SendNotification.Sent = True
        SendNotification.Trash = True
        Resume Next
    End If
End Function
Public Function GetLog(strFilePath As String) As String
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Set ts = fso.OpenTextFile(strFilePath)
    GetLog = ts.ReadAll
End Function
Public Sub Logger(Message As String)
    Dim tmpMsg     As String
    Dim strLogPath As String
    If Dir$(strLogLoc, vbDirectory) = "" Then MkDir strLogLoc 'Environ$("APPDATA") & "\JPTRS\"
    strLogPath = strLogLoc & "\LOG.LOG"
    'JPTRS.rtbLog.SelText = strLogBuffer
    Open strLogPath For Append As #1
    With JPTRS
        tmpMsg = DateTime.Date & " " & DateTime.Time & ": " & Message
        .rtbLog.SelText = tmpMsg & vbNewLine
        strLogBuffer = strLogBuffer + tmpMsg & vbNewLine
        Print #1, tmpMsg
        Close #1
        ' DoEvents
    End With
    If JPTRS.TCPServer.State = 7 And Not bolWaitingForPass Then
        SocketLog tmpMsg, LogPacket
    End If
End Sub
Public Sub ClearLog()
    Kill strLogLoc & "\LOG.LOG"
    strLogBuffer = ""
    Logger "Log Cleared..."
End Sub
Public Sub Wait(ByVal DurationMS As Long)
    Dim EndTime As Long
    EndTime = GetTickCount + DurationMS
    Logger "Waiting... " & DurationMS & "ms"
    Do While EndTime > GetTickCount
        DoEvents
        Sleep 1
    Loop
    Do Until Not bolExecutionPaused
        DoEvents
        Sleep 1
    Loop
End Sub
