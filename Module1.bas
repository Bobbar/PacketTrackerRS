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
Public Const strSMTPServer    As String = "" '"mx.wthg.com"
Public Const strServerAddress As String = "localhost" '"ohbre-pwadmin01"
Public Const strUsername      As String = "TicketApp"
Public Const strPassword      As String = "yb4w4"
Public Const strListenPort As String = "1001"
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
Public Type NotificationReturnType
    Sent As Boolean
    Trash As Boolean
End Type
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
    Status As NotificationReturnType
    TimeStamp As String
    'Trash As Boolean
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
Public ReportData() As ReportInfo
Public EmailData()  As EmailInfo
Public bolVerbose   As Boolean
Public strLogLoc    As String
Public strCSVLoc    As String
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
Public Const DayToRun                As Long = vbMonday
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
Public dtRunDate             As Date
Public intRetryFail          As Integer
Public Const intRetryFailMax As Integer = 5
Public Const Weekly          As Integer = 0
Public Const Daily           As Integer = 1
Public Const TimeToRun       As Integer = 12 'daily report run hour
Public Const Controls        As String = "CONTROLS"
Public Const IM              As String = "INDUSTRIAL MACH"
Public Const Nuclear         As String = "NUCLEAR"
Public Const RockyMtn        As String = "ROCKY MT"
Public Const SteelFab        As String = "STEEL FAB"
Public Const Wooster         As String = "WOOSTER"
Private Type UserAttributes
    UserName As String
    FullName As String
    EMail As String
    GetsDaily As Boolean
    GetsWeekly As Boolean
    Filters As String
End Type
Private Users() As UserAttributes
Private Type GroupAttribs
    Filters As String
    Userlist() As UserAttributes
End Type
Private ReportsGroups() As GroupAttribs


Public Sub RefreshUserList()
    With JPTRS
        .tmrCheckQueue.Enabled = False
        .tmrReportClock.Enabled = False
        GetUserIndex
        .tmrCheckQueue.Enabled = True
        .tmrReportClock.Enabled = True
    End With
End Sub
Public Function FilterString(Filter As String) As String
    Dim FilterArray() As String
    Dim tmpString     As String
    FilterArray = Split(Filter, ",")
    If FilterArray(0) = 1 Then tmpString = tmpString + "| Controls |"
    If FilterArray(1) = 1 Then tmpString = tmpString + "| Industrial Machine |"
    If FilterArray(2) = 1 Then tmpString = tmpString + "| Nuclear |"
    If FilterArray(3) = 1 Then tmpString = tmpString + "| Rocky Mtn |"
    If FilterArray(4) = 1 Then tmpString = tmpString + "| Steel Fab |"
    If FilterArray(5) = 1 Then tmpString = tmpString + "| Wooster |"
    FilterString = tmpString
End Function
Public Function IsUnFiltered(Filters As String, _
                             Plant As String) As Boolean 'is the plant filtered?
    Dim FilterArray() As String
    Dim IsFirst       As Boolean
    Dim i             As Integer
    Dim strSQL        As String
    IsUnFiltered = False
    FilterArray = Split(Filters, ",")
    If FilterArray(0) = 1 And Plant = Controls Then IsUnFiltered = True
    If FilterArray(1) = 1 And Plant = IM Then IsUnFiltered = True
    If FilterArray(2) = 1 And Plant = Nuclear Then IsUnFiltered = True
    If FilterArray(3) = 1 And Plant = RockyMtn Then IsUnFiltered = True
    If FilterArray(4) = 1 And Plant = SteelFab Then IsUnFiltered = True
    If FilterArray(5) = 1 And Plant = Wooster Then IsUnFiltered = True
End Function
Public Sub GetReportGroups() 'create a list of report groups with list a users included
    Dim a As Integer, b As Integer
    ReDim ReportsGroups(0) As GroupAttribs
    For a = 1 To UBound(Users)
        If Users(a).GetsDaily Then
            If UBound(ReportsGroups) < 1 Then
                NewGroup ReportsGroups, Users, a
            Else
                If Not InGroupIndex(ReportsGroups, Users(a).Filters) Then
                    NewGroup ReportsGroups, Users, a
                ElseIf InGroupIndex(ReportsGroups, Users(a).Filters) Then
                    AddToGroup ReportsGroups, Users, GroupIndex(ReportsGroups, Users(a).Filters), a
                End If
            End If
        End If
    Next a
End Sub
Private Sub NewGroup(GroupIndex() As GroupAttribs, _
                     UserIndex() As UserAttributes, _
                     Index As Integer) 'create a new report group add user to it
    ReDim Preserve ReportsGroups(UBound(ReportsGroups) + 1)
    ReportsGroups(UBound(ReportsGroups)).Filters = UserIndex(Index).Filters
    ReDim ReportsGroups(UBound(ReportsGroups)).Userlist(0)
    ReportsGroups(UBound(ReportsGroups)).Userlist(0) = UserIndex(Index)
End Sub
Private Sub AddToGroup(GroupIndex() As GroupAttribs, _
                       UserIndex() As UserAttributes, _
                       Index As Integer, _
                       UIndex As Integer) 'add user to existing group
    With GroupIndex(Index)
        ReDim Preserve .Userlist(UBound(.Userlist) + 1)
        .Userlist(UBound(.Userlist)) = UserIndex(UIndex)
    End With
End Sub
Private Function InGroupIndex(GroupArray() As GroupAttribs, _
                              Filters As String) As Boolean 'is there already a group for this filter?
    Dim i As Long
    InGroupIndex = False
    For i = 0 To UBound(GroupArray)
        If Filters = GroupArray(i).Filters Then
            InGroupIndex = True
            Exit For
        End If
    Next i
End Function
Private Function GroupIndex(GroupArray() As GroupAttribs, _
                            Filters As String) As Integer 'where is the existing filter?
    Dim i As Long
    GroupIndex = 0
    For i = 0 To UBound(GroupArray)
        If Filters = GroupArray(i).Filters Then
            GroupIndex = i
            Exit For
        End If
    Next i
End Function
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
    If bolVerbose Then Logger "Clearing Queue..."
    For i = 0 To UBound(EmailData) - 1
        strSQL1 = "SELECT * From emailqueue Where idGUID = '" & EmailData(i).GUID & "'"
        JPTRS.lblStatus.Caption = "Clearing Queue..."
        rs.Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
        rs.Delete
        rs.Update
        rs.Close
    Next i
    ReDim EmailData(0)
    Exit Sub
errs:
    Logger "Error Clearing Queue!"
    Logger "ERROR DTL:  SUB = ClearEmailQueue | " & Err.Number & " - " & Err.Description
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
        tmpHTML = tmpHTML + "<FONT SIZE=2>" & strTimeStamp & "</FONT><BR>" & vbCrLf
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
    Logger "Failed to Genterate HTML!"
    If bolVerbose Then Logger "ERROR DTL:  SUB = GenerateHTML | " & Err.Number & " - " & Err.Description
End Function
Public Function GetEmail(strUsername As String) As String
    Dim i As Integer
    For i = 0 To UBound(Users)
        If Users(i).UserName = strUsername Then
            GetEmail = UCase$(Users(i).EMail)
            Exit Function
        End If
    Next i
End Function
Public Function GetFullName(strUsername As String) As String
    Dim i As Integer
    For i = 0 To UBound(Users)
        If Users(i).UserName = strUsername Then
            GetFullName = UCase$(Users(i).FullName)
            Exit Function
        End If
    Next i
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
Public Function TimeForDaily() As Boolean
    Dim Flag As Boolean
    Flag = CBool(GetSetting(App.EXEName, "DailyReportSent", "Sent", 0))
    TimeForDaily = False
    If IsWeekday Then
        If Hour(Now) = TimeToRun And Flag Then
            TimeForDaily = False
        ElseIf Hour(Now) <> TimeToRun And Flag Then
            TimeForDaily = False
            SaveSetting App.EXEName, "DailyReportSent", "Sent", "FALSE"
        ElseIf Hour(Now) <> TimeToRun And Not Flag Then
            TimeForDaily = False
        ElseIf Hour(Now) = TimeToRun And Not Flag Then
            TimeForDaily = True
        End If
        'JPTRS.lblReportStatus.Caption = Str(Flag)
    End If
End Function
Public Function IsWeekday() As Boolean
    IsWeekday = False
    If Weekday(Now) = 1 Or Weekday(Now) = 7 Then
        IsWeekday = False
    ElseIf Weekday(Now) <> 1 Or Weekday(Now) <> 7 Then
        IsWeekday = True
    End If
End Function
Public Function WeeklyReportRecpts() As String
    WeeklyReportRecpts = ""
    Dim tmpRecpts As String
    Dim i         As Integer
    For i = 0 To UBound(Users)
        If Users(i).GetsWeekly Then tmpRecpts = tmpRecpts + Users(i).EMail + ";"
    Next i
    WeeklyReportRecpts = tmpRecpts
End Function
Public Function DailyReportRecpts(GroupIndex As Integer) As String
    DailyReportRecpts = ""
    Dim tmpRecpts As String
    Dim i         As Integer
    For i = 0 To UBound(ReportsGroups(GroupIndex).Userlist)
        tmpRecpts = tmpRecpts + ReportsGroups(GroupIndex).Userlist(i).EMail + ";"
    Next i
    DailyReportRecpts = tmpRecpts
End Function
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
Public Sub SendReport(intReportType As Integer, MailFrom As String, MailTo As String)
    If intReportType = Weekly Then
        Logger "Sending Weekly Report...  (" & dtStartDate & " to " & dtEndDate & ")"
    ElseIf intReportType = Daily Then
        Logger "Sending Daily Report... "
    End If
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
        If intReportType = Weekly Then
            .Subject = "JPT: Weekly Report (" & dtStartDate & " to " & dtEndDate & ")"
        ElseIf intReportType = Daily Then
            .Subject = "JPT: Daily Report (" & dtEndDate & ")"
        End If
        .HTMLBody = strReportHTML
        '.TextBody = Message
        .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    'Set registry flag telling me that the report for this week has been sent
    If intReportType = Weekly Then
        SaveSetting App.EXEName, "WeeklyReportSent", "Sent", "TRUE"
    ElseIf intReportType = Daily Then
        SaveSetting App.EXEName, "DailyReportSent", "Sent", "TRUE"
    End If
    Logger "Report Sent..."
    Exit Sub
errs:
    If intReportType = Weekly Then
        Logger "***** Failed to send Weekly Report! *****"
    ElseIf intReportType = Daily Then
        Logger "***** Failed to send Daily Report! *****"
    End If
    If bolVerbose Then Logger "***** ERROR DTL:  SUB = SendReport | " & Err.Number & " - " & Err.Description
    Logger "Waiting and retrying..."
    'Logger "Waiting..."
    Wait intWaitTime
    Logger "Retrying..."
    SendReport intReportType, MailFrom, MailTo
End Sub
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
Public Sub Logger(Message As String)
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
    If JPTRS.TCPServer.State = 7 Then
        SocketLog Message
    End If

End Sub
Public Sub Wait(ByVal DurationMS As Long)
    Dim EndTime As Long
    EndTime = GetTickCount + DurationMS
    Logger "Waiting... " & DurationMS & "ms"
    Do While EndTime > GetTickCount
        DoEvents
        Sleep 1
    Loop
End Sub
Public Sub WeeklyReportGetData()
    On Error GoTo errs
    Dim lngTIS  As Long
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    i = 0
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM packetlist d LEFT JOIN packetentrydb c ON c.idJobNum = d.idJobNum WHERE idDate = (SELECT MAX(idDate) FROM packetentrydb c2 Where c2.idJobNum = d.idJobNum) AND idStatus='OPEN' ORDER BY idDate DESC LIMIT 50"
    Logger "Starting Weekly Report..."
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
                    ReportData(i).Action = !idAction & " by " & GetFullName(!idUser)
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
    ReportParseHTML Weekly
    SendReport Weekly, "JPTReportServer@worthingtonindustries.com", WeeklyReportRecpts
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "WeeklyReportGetData"
End Sub
Public Sub RunDailyReport()
    On Error GoTo errs
    Dim i As Integer
    Logger "Starting Daily Report..."
    Logger "Processing filter groups..."
    GetReportGroups
    Logger UBound(ReportsGroups) & " groups found..."
    Logger "Precessing reports..."
    For i = 1 To UBound(ReportsGroups)
        Logger "Processing report " & i & " of " & UBound(ReportsGroups) & "..."
        DailyReportGetData ReportsGroups(i).Filters
        Logger "Recipients: " & DailyReportRecpts(i)
        SendReport Daily, "JPTReportServer@worthingtonindustries.com", DailyReportRecpts(i)
        'Logger "Waiting..."
        Wait intWaitTime
    Next i
    Logger "Daily report complete!"
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "RunDailyReport"
End Sub
Public Sub DailyReportGetData(Filters As String)
    On Error GoTo errs
    Dim lngTIS  As Long
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    i = 0
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM packetlist d LEFT JOIN packetentrydb c ON c.idJobNum = d.idJobNum WHERE idDate = (SELECT MAX(idDate) FROM packetentrydb c2 Where c2.idJobNum = d.idJobNum) AND idStatus='OPEN' ORDER BY idDate DESC LIMIT 50"
    Logger "Starting Daily Report... Filters = " & FilterString(Filters)
    Set rs = cn_global.Execute(strSQL1)
    dtStartDate = DateTime.Date
    dtEndDate = DateTime.Date
    If rs.RecordCount < 1 Then Exit Sub
    ReDim ReportData(0)
    With rs
        Do Until .EOF
            If Format(!idDate, strUserDateFormat) >= dtStartDate And Format(!idDate, strUserDateFormat) <= dtEndDate And IsUnFiltered(Filters, !idPlant) Then
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
                    ReportData(i).Action = !idAction & " by " & GetFullName(!idUser)
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
    ReportParseHTML Daily, Filters
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "DailyReportGetData"
End Sub
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
    Logger "Failed while parsing Weekly Report!"
    If bolVerbose Then Logger "ERROR DTL:  SUB = WeeklyReportParseCSV | " & Err.Number & " - " & Err.Description
End Sub
Public Sub ReportParseHTML(intReportType As Integer, _
                           Optional Filters As String) '0 = Weekly 1 = Daily
    On Error GoTo errs
    Logger "Parsing report to HTML..."
    Dim tmpHTML As String
    Dim i       As Integer
    tmpHTML = tmpHTML + "<FONT STYLE=font-family:Tahoma;>" & vbCrLf
    If intReportType = Weekly Then
        tmpHTML = tmpHTML + "<FONT SIZE=4><U>Weekly Job Packet Report</U></FONT><BR>" & vbCrLf
    ElseIf intReportType = Daily Then
        tmpHTML = tmpHTML + "<FONT SIZE=4><U>Daily Job Packet Report</U></FONT><BR>" & vbCrLf
        tmpHTML = tmpHTML + "<FONT SIZE=2> Filters:  " & FilterString(Filters) & "</FONT><BR>" & vbCrLf
    End If
    tmpHTML = tmpHTML + "<FONT SIZE=2> Run Date: " & Now & "</FONT><BR>" & vbCrLf
    If intReportType = Weekly Then
        tmpHTML = tmpHTML + "<FONT SIZE=2> Report Date: " & dtStartDate & " to " & dtEndDate & "</FONT><BR>" & vbCrLf
    End If
    tmpHTML = tmpHTML + "<BR>" & vbCrLf
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
    If intReportType = Weekly Then
        Logger "Failed while parsing Weekly Report HTML!"
    ElseIf intReportType = Daily Then
        Logger "Failed while parsing Daily Report HTML!"
    End If
    If bolVerbose Then Logger "ERROR DTL:  SUB =  ReportParseHTML | " & Err.Number & " - " & Err.Description
End Sub
