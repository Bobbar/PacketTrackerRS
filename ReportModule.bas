Attribute VB_Name = "ReportModule"
Public Const colCreate    As Long = &HFFBF80 '&H80C0FF
Public Const colInTransit As Long = &H80FF80
Public Const colReceived  As Long = &HF4FF80 '&H80FFFF
Public Const colClosed    As Long = &H8080FF
Public Const colFiled     As Long = &H8587FF '&HFF8080
Public Const colReopened  As Long = &HFF80FF
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
Public ReportData()    As ReportInfo
Public EmailData()     As EmailInfo
Public strReportHTML   As String
Public dtStartDate     As Date, dtEndDate As Date
Public Const DayToRun  As Long = vbMonday
Public dtRunDate       As Date
Public Const Weekly    As Integer = 0
Public Const Daily     As Integer = 1
Public Const TimeToRun As Integer = 12 'daily report run hour
Public Const Controls  As String = "CONTROLS"
Public Const IM        As String = "INDUSTRIAL MACH"
Public Const Nuclear   As String = "NUCLEAR"
Public Const RockyMtn  As String = "ROCKY MT"
Public Const SteelFab  As String = "STEEL FAB"
Public Const Wooster   As String = "WOOSTER"
Private Type ReportAttributes
    Name As String
    ID As Integer
    RunDay As Integer
    RunTime As Integer 'String
    StartDate As Date
    EndDate As Date
    CompanyFilter As String
    HasRun As Boolean
End Type
Public Reports() As ReportAttributes
Private Type ReportGroupAttributes
    GroupID As Integer
    ReportID As Integer
    EntryID As Integer
End Type
Private ReportGroups() As ReportGroupAttributes
Private Function GetReportRecipients(ReportID As Integer) As String
    On Error GoTo errs
    Dim rs            As New ADODB.Recordset
    Dim strSQL1       As String
    Dim i             As Integer
    Dim tmpRecipients As String
    strSQL1 = "SELECT users_0.idFullName, users_0.idEmail, users_0.idGroupID, reportsgroups_0.idGroupID, reportsgroups_0.idReportID, reportsgroups_0.idEntryID, reports_0.idReportName" & " FROM ticketdb.reports reports_0, ticketdb.reportsgroups reportsgroups_0, ticketdb.users users_0" & " WHERE users_0.idGroupID = reportsgroups_0.idGroupID AND reports_0.idReportID = reportsgroups_0.idReportID AND ((reportsgroups_0.idReportID='" & ReportID & "'))"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset
    If rs.RecordCount = 0 Then
        GetReportRecipients = ""
        Exit Function
    End If
    Do Until rs.EOF
        With rs
            tmpRecipients = tmpRecipients + !idEmail & ";"
            .MoveNext
        End With
    Loop
    GetReportRecipients = tmpRecipients
    Exit Function
errs:
    ErrHandle Err.Number, Err.Description, "GetReportRecipients"
End Function
Public Sub GetReports()
    On Error GoTo errs
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim i       As Integer
    strSQL1 = "SELECT * FROM reports"
    cn_global.CursorLocation = adUseClient
    rs.Open strSQL1, cn_global, adOpenKeyset
    i = 1
    ReDim Reports(rs.RecordCount)
    Do Until rs.EOF
        With rs
            Reports(.AbsolutePosition).Name = !idReportName
            Reports(.AbsolutePosition).ID = !idReportID
            Reports(.AbsolutePosition).RunDay = !idRunDay
            Reports(.AbsolutePosition).RunTime = CInt(!idRunTime)
            Reports(.AbsolutePosition).StartDate = ParseDate(!idStartDate)
            Reports(.AbsolutePosition).EndDate = ParseDate(!idEndDate)
            Reports(.AbsolutePosition).CompanyFilter = !idCompanyFilter
            Reports(.AbsolutePosition).HasRun = CBool(!idHasRun)
            rs.MoveNext
        End With
    Loop
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "GetReports"
End Sub
Private Function ParseDate(DateCode As String) As Date
    Dim strVar    As String
    Dim Modifier  As Integer
    Dim SplitCode As Variant
    Dim tmpDate   As Date
    ParseDate = DateTime.Date
    SplitCode = Split(DateCode, "|")
    If UBound(SplitCode) > 0 Then
        strVar = SplitCode(0)
        Modifier = SplitCode(1)
        Select Case strVar
            Case "TODAY"
                tmpDate = DateAdd("d", Modifier, DateTime.Date)
        End Select
    ElseIf UBound(SplitCode) = 0 Then
        strVar = SplitCode(0)
        tmpDate = DateTime.Date
    End If
    ParseDate = tmpDate
End Function
Public Sub RunReports()
    Dim i As Integer
    For i = 0 To UBound(Reports)
        If IsDay(Reports(i).RunDay) Then
            If Hour(Now) = Reports(i).RunTime And Reports(i).HasRun Then
                'do nothing
            ElseIf Hour(Now) <> Reports(i).RunTime And Reports(i).HasRun Then
                UpdateReportHasRun Reports(i), False
                GetReports
            ElseIf Hour(Now) <> Reports(i).RunTime And Not Reports(i).HasRun Then
                'do nothing
            ElseIf Hour(Now) = Reports(i).RunTime And Not Reports(i).HasRun Then
                StartReport Reports(i)
            End If
        ElseIf Reports(i).RunDay = 8 And IsWeekday Then '8 = every weekday
            If Hour(Now) = Reports(i).RunTime And Reports(i).HasRun Then
                'do nothing
            ElseIf Hour(Now) <> Reports(i).RunTime And Reports(i).HasRun Then
                UpdateReportHasRun Reports(i), False
                GetReports
            ElseIf Hour(Now) <> Reports(i).RunTime And Not Reports(i).HasRun Then
                'do nothing
            ElseIf Hour(Now) = Reports(i).RunTime And Not Reports(i).HasRun Then
                StartReport Reports(i)
            End If
        End If
    Next i
End Sub
Public Sub UpdateReportHasRun(Report As ReportAttributes, HasRun As Boolean)
    On Error GoTo errs
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * From reports Where idReportName = '" & Report.Name & "' AND idReportID = '" & Report.ID & "'"
    With rs
        .Open strSQL1, cn_global, adOpenKeyset, adLockOptimistic
        !idHasRun = CInt(Int(HasRun))
        .Update
    End With
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "UpdateReportHasRun"
End Sub
Public Sub StartReport(Report As ReportAttributes)
    Logger "Starting Report: " & Report.Name & "..."
    GetReportData Report
    Logger "Report Complete..."
End Sub
Public Sub GetReportData(Report As ReportAttributes)
    On Error GoTo errs
    Dim lngTIS  As Long
    Dim i       As Integer
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    i = 0
    cn_global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM packetlist d LEFT JOIN packetentrydb c ON c.idJobNum = d.idJobNum WHERE idDate = (SELECT MAX(idDate) FROM packetentrydb c2 Where c2.idJobNum = d.idJobNum) AND idStatus='OPEN' ORDER BY idDate DESC LIMIT 50"
    dtStartDate = Report.StartDate 'dateTime.Date
    dtEndDate = Report.EndDate 'DateTime.Date
    Logger "Getting Report Data..." & vbCrLf & "Filters = " & FilterString(Report.CompanyFilter) & vbCrLf & "StartDate: " & dtStartDate & vbCrLf & "EndDate: " & dtEndDate
    Set rs = cn_global.Execute(strSQL1)
    If rs.RecordCount < 1 Then Exit Sub
    ReDim ReportData(0)
    With rs
        Do Until .EOF
            If Format(!idDate, strUserDateFormat) >= dtStartDate And Format(!idDate, strUserDateFormat) <= dtEndDate And IsUnFiltered(Report.CompanyFilter, !idPlant) Then
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
    ParseReportHTML Report
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "DailyReportGetData"
End Sub
Public Sub ParseReportHTML(Report As ReportAttributes)  '0 = Weekly 1 = Daily
    On Error GoTo errs
    Logger "Parsing Report to HTML..."
    Dim tmpHTML As String
    Dim i       As Integer
    tmpHTML = tmpHTML + "<FONT STYLE=font-family:Tahoma;>" & vbCrLf
    tmpHTML = tmpHTML + "<FONT SIZE=4><U>Job Packet Report: " & Report.Name & "</U></FONT><BR>" & vbCrLf
    tmpHTML = tmpHTML + "<FONT SIZE=2> Filters:  " & FilterString(Report.CompanyFilter) & "</FONT><BR>" & vbCrLf
    tmpHTML = tmpHTML + "<FONT SIZE=2> Run Date: " & Now & "</FONT><BR>" & vbCrLf
    If intReportType = Weekly Then
        tmpHTML = tmpHTML + "<FONT SIZE=2> Report Date: " & Report.StartDate & " to " & Report.EndDate & "</FONT><BR>" & vbCrLf
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
    SendReport Report
    Exit Sub
errs:
    Logger "Failed while parsing Report HTML!  ReportName: " & Report.Name
    If bolVerbose Then Logger "ERROR DTL:  SUB =  ReportParseHTML | " & Err.Number & " - " & Err.Description
End Sub
Public Sub SendReport(Report As ReportAttributes)
    On Error GoTo errs
    Dim iConf As New CDO.Configuration
    Dim Flds  As ADODB.Fields
    Dim iMsg  As New CDO.Message
    Set Flds = iConf.Fields
    Dim MailTo As String
    MailTo = GetReportRecipients(Report.ID)
    If MailTo = "" Then
        Logger "No recipients found for this report. Moving on..."
        UpdateReportHasRun Report, True
        Exit Sub
    End If
    Logger "Sending Report...  (" & Report.Name & " to " & MailTo & ")"
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
        .From = "JPTReportServer@worthingtonindustries.com" 'MailFrom 'GetEmail(MailFrom)
        .Subject = "JPT Report: " & Report.Name
        '.Subject = "JPT: Daily Report (" & dtEndDate & ")"
        .HTMLBody = strReportHTML
        '.TextBody = Message
        .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    'Set registry flag telling me that the report for this week has been sent
    UpdateReportHasRun Report, True
    Logger "Report Sent..."
    strReportHTML = ""
    Wait intWaitTime
    Exit Sub
errs:
    Logger "***** Failed to send Report! ReportID: " & Report.ID & " *****"
    If bolVerbose Then Logger "***** ERROR DTL:  SUB = SendReport | " & Err.Number & " - " & Err.Description
    Logger "Waiting and retrying..."
    'Logger "Waiting..."
    Wait intWaitTime
    Logger "Retrying..."
    SendReport Report
End Sub
Public Function IsDay(intDay As Integer) As Boolean
    IsDay = False
    If Weekday(Now) = intDay Then
        IsDay = True
    Else
        IsDay = False
    End If
End Function
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
        Case 8
            strDayOfWeek = "EveryWeekDay"
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
    Logger "Failed while parsing Weekly Report!"
    If bolVerbose Then Logger "ERROR DTL:  SUB = WeeklyReportParseCSV | " & Err.Number & " - " & Err.Description
End Sub
