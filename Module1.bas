Attribute VB_Name = "Module1"
Global cn_global As New ADODB.Connection
Public Const colCreate       As Long = &H80C0FF
Public Const colInTransit    As Long = &H80FF80
Public Const colReceived     As Long = &H80FFFF
Public Const colClosed       As Long = &H8080FF
Public Const colFiled        As Long = &HFF8080
Public Const colReopened     As Long = &HFF80FF
Public Const strServerAddress As String = "10.35.1.40"
Public Const strUsername As String = "TicketApp"
Public Const strPassword As String = "yb4w4"
Public strSQLDriver As String
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
Public Sub FindMySQLDriver()
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
                            strCreateDate As String, strComment As String)
    'On Error GoTo errs
    Dim iConf As New CDO.Configuration
    Dim Flds  As ADODB.Fields
    Dim iMsg  As New CDO.Message
    Set Flds = iConf.Fields
    ' Set the configuration
    Flds(cdoSendUsingMethod) = cdoSendUsingPort
    Flds(cdoSMTPServer) = "mx.wthg.com"
    ' ... other settings
    Flds.Update
    With iMsg
        Set .Configuration = iConf
        .Sender = GetEmail(MailFrom)
        .To = GetEmail(MailTo)
        .From = GetEmail(MailFrom)
        If UCase$(SendRec) = "SEND" Then
            .Subject = "JPT: " & GetFullName(strLocalUser) & " sent you a packet"
        ElseIf UCase$(SendRec) = "REC" Then
            .Subject = "JPT: " & GetFullName(strLocalUser) & " received a packet"
        End If
        .HTMLBody = GenerateHTML(SendRec, GetFullName(MailFrom), MailTo, JobNum, strDescrip, strPartNum, strCustomer, strCreator, strCreateDate, strComment)
        '.TextBody = Message
        .Send
    End With
    Set iMsg = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    'Exit Sub
    'errs:
    ' Debug.Print Err.Number
    ' If Err.Number = -2147220973 Then
    '    MsgBox "Failed to send EMail notification!"
    '  End If
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
                             strComment As String) As String
    ' On Error GoTo errs
    Dim tmpHTML As String
    If UCase$(strSendOrRec) = "SEND" Then
        Dim BackColor As String
        BackColor = Hex$(colInTransit)
        tmpHTML = tmpHTML + "<HTML>" & vbCrLf
        tmpHTML = tmpHTML + "<BODY BGCOLOR=" & BackColor & ">" & vbCrLf
        tmpHTML = tmpHTML + "<FONT STYLE=font-family:Tahoma;>" & vbCrLf
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
        tmpHTML = tmpHTML + "<td>" & strCreator & "</td>" & vbCrLf
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
        tmpHTML = tmpHTML + "<td>" & strCreator & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "<td>" & strCreateDate & "</td>" & vbCrLf
        tmpHTML = tmpHTML + "</tr>" & vbCrLf
        tmpHTML = tmpHTML + "</table>" & vbCrLf
        tmpHTML = tmpHTML + " <FONT>" & vbCrLf
        tmpHTML = tmpHTML + " </BODY>" & vbCrLf
        tmpHTML = tmpHTML + " </HTML>" & vbCrLf
        GenerateHTML = tmpHTML
    End If
    '  Exit Function
    'errs:
    '    Debug.Print Err.Number
End Function

