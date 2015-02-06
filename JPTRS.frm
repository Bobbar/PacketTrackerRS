VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form JPTRS 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JPTRS"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "JPTRS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock TCPServer 
      Left            =   7080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrTaskTimer 
      Interval        =   60000
      Left            =   7560
      Top             =   1080
   End
   Begin VB.Timer tmrReportClock 
      Interval        =   3000
      Left            =   7560
      Top             =   600
   End
   Begin VB.CommandButton cmdSendToTray 
      Caption         =   "Minimize To Tray"
      Height          =   480
      Left            =   7020
      TabIndex        =   4
      Top             =   1620
      Width           =   990
   End
   Begin VB.CheckBox chkVerbose 
      Caption         =   "Verbose"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   4920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.ListBox lstLog 
      Height          =   2595
      Left            =   180
      TabIndex        =   0
      Top             =   2220
      Width           =   7995
   End
   Begin VB.Timer tmrCheckQueue 
      Interval        =   2000
      Left            =   7560
      Top             =   120
   End
   Begin VB.Label lblRetries 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%RETRIES%"
      Height          =   195
      Left            =   1920
      TabIndex        =   23
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label lblSuccess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%SUCCESS%"
      Height          =   195
      Left            =   1920
      TabIndex        =   22
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label lblRequests 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%REQUESTS%"
      Height          =   195
      Left            =   1920
      TabIndex        =   21
      Top             =   1080
      Width           =   1110
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retries:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   20
      Top             =   1560
      Width           =   660
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Successful:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   900
      TabIndex        =   19
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Requests: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1020
      TabIndex        =   18
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label lblSMTPAddy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%SMTP Addy%"
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   600
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   16
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label lblServerIP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%SERVER IP%"
      Height          =   195
      Left            =   1680
      TabIndex        =   15
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   14
      Top             =   360
      Width           =   840
   End
   Begin VB.Label lblCurDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%CUR DAY%"
      Height          =   195
      Left            =   6240
      TabIndex        =   13
      Top             =   960
      Width           =   990
   End
   Begin VB.Label lblCurDayLBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Day:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   12
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label lblRptDay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%RPT DAY%"
      Height          =   195
      Left            =   6240
      TabIndex        =   11
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblDayLBL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Day:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5100
      TabIndex        =   10
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label lblReportStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%FLAG%"
      Height          =   195
      Left            =   6180
      TabIndex        =   9
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblReportSent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report Sent:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   8
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      Height          =   195
      Left            =   3840
      TabIndex        =   7
      Top             =   1860
      Width           =   330
   End
   Begin VB.Label lblCodedBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coded by: Bobby Lovell"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   6660
      TabIndex        =   6
      Top             =   5040
      Width           =   1515
   End
   Begin VB.Label lblAPPVERSION 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%APP VERSION%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   3420
      TabIndex        =   5
      Top             =   5040
      Width           =   1290
   End
   Begin VB.Label lblStatusLBL 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Left            =   3480
      TabIndex        =   3
      Top             =   1560
      Width           =   525
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "_________"
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   1560
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   3360
      Picture         =   "JPTRS.frx":08CA
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "JPTRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nid As NOTIFYICONDATA
Private Function CurrentInterval(lngTimeCounted As Long, _
                                 lngIntervalTime As Long) As Boolean
    CurrentInterval = (lngTimeCounted / lngIntervalTime) = Int(lngTimeCounted / lngIntervalTime)
End Function
Sub minimize_to_tray()
    Me.Hide
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon ' the icon will be your Form1 project icon
    nid.szTip = "Click to View" & vbNullChar
    Shell_NotifyIcon NIM_ADD, nid
End Sub
Private Sub chkVerbose_Click()
    bolVerbose = CBool(chkVerbose.Value)
End Sub
Private Sub cmdSendToTray_Click()
    minimize_to_tray
End Sub
Private Sub Form_Initialize()
    dtRunDate = DateTime.Now
End Sub
Private Sub Form_Load()
    minimize_to_tray
    DoEvents
    strAPPTITLE = JPTRS.Caption
    lblServerIP.Caption = GetIpAddrTable(0)
    lblSMTPAddy.Caption = strSMTPServer
    lblAPPVERSION.Caption = App.Major & "." & App.Minor & "." & App.Revision
    bolVerbose = CBool(chkVerbose.Value)
    strLogLoc = App.Path  'Environ$("APPDATA") & "\JPTRS\LOG.LOG"
    strCSVLoc = Environ$("APPDATA") & "\JPTRS\"
    bolExecutionPaused = False
    Logger "Initializing..."
    FindMySQLDriver
    Logger "Starting Global ADO Connection..."
    If bolVerbose Then Logger "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    If ConnectToDB Then Logger "Connected!"
    'cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    Logger "Getting User List..."
    GetUserIndex
    Logger "Starting TCP Server..."
    StartTCPServer
    Logger "Ready!..."
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim msg     As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONDOWN
            Me.Show ' show form
            Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
        Case WM_RBUTTONDOWN
        Case WM_RBUTTONUP
            Me.Show
            Shell_NotifyIcon NIM_DELETE, nid
        Case WM_RBUTTONDBLCLK
    End Select
End Sub
Public Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    EndProgram
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
End Sub
Private Sub TCPServer_Close()
    Logger "TCP Socket: " & strSocketAcceptedID & " disconnecting..."
    TCPServer.Close
    strSocketRequestID = ""
    strSocketAcceptedID = ""
    StartTCPServer
End Sub
Private Sub TCPServer_ConnectionRequest(ByVal requestID As Long)
    If TCPServer.State <> sckClosed Then TCPServer.Close
    ' Accept the request with the requestID
    ' parameter.
    TCPServer.Accept requestID
    strSocketRequestID = requestID
    Logger "TCP Socket: Connection attempt from " & requestID
    Logger "TCP Socket: Requesting Auth"
    RequestPass
End Sub
Private Sub TCPServer_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    TCPServer.GetData strData
    ParsePacket strData
End Sub
Private Sub tmrCheckQueue_Timer()
    On Error GoTo errs
    JPTRS.lblStatus.Caption = "Idle..."
    CheckQueue
    'JPTRS.Caption = strAPPTITLE + " - Up " & ConvertTime(DateTime.Now)
    lblRequests.Caption = lngAttempts
    lblSuccess.Caption = lngSuccess
    lblRetries.Caption = lngRetries
    'DoEvents
    Exit Sub
errs:
    ErrHandle Err.Number, Err.Description, "CheckQueueTimer" 'Logger Err.Number & " - " & Err.Description
    Resume Next
End Sub
Private Sub tmrReportClock_Timer()
    lblTime = Now
    If OKToRun Then WeeklyReportGetData
    If TimeForDaily Then
        RunDailyReport
    End If
    lblRptDay.Caption = strDayOfWeek(DayToRun)
    lblCurDay.Caption = strDayOfWeek(Weekday(Now))
End Sub
Private Sub tmrTaskTimer_Timer()
    MinsCounted = MinsCounted + 1
    If CurrentInterval(MinsCounted, MinutesTillRefresh) Then
        RefreshUserList
    Else
    End If
    If CurrentInterval(MinsCounted, MinutesTillStatusReport) Then
        Logger "STATUS: Uptime: " & ConvertTime(DateTime.Now) & "    Atmpts, Sucss, Rtry: " & lngAttempts & ", " & lngSuccess & ", " & lngRetries
    End If
    If MinsCounted >= 1440 Then MinsCounted = 0 'day has passed, start over
End Sub
