VERSION 5.00
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
   Begin VB.CommandButton cmdSendToTray 
      Caption         =   "Minimize To Tray"
      Height          =   480
      Left            =   7020
      TabIndex        =   4
      Top             =   1560
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
      Top             =   240
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
      Left            =   3300
      TabIndex        =   3
      Top             =   1800
      Width           =   525
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "_________"
      Height          =   195
      Left            =   3900
      TabIndex        =   2
      Top             =   1800
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   3360
      Picture         =   "JPTRS.frx":08CA
      Top             =   300
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
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim msg     As Long
    Dim sFilter As String
    msg = x / Screen.TwipsPerPixelX
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
Private Sub chkVerbose_Click()
    bolVerbose = CBool(chkVerbose.Value)
End Sub

Private Sub cmdSendToTray_Click()
    minimize_to_tray
End Sub

Private Sub Form_Load()
    lblAPPVERSION.Caption = App.Major & "." & App.Minor & "." & App.Revision
    bolVerbose = CBool(chkVerbose.Value)
    strLogLoc = Environ$("APPDATA") & "\JPTRS\LOG.LOG"
    FindMySQLDriver
    cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
    GetUserIndex
    minimize_to_tray
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim blah
    blah = MsgBox("Are you sure you want to close the server!", vbOKCancel, "Are you sure?!")
    If blah = vbOK Then
        cn_global.Close
        Unload Me
        End
    Else
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
End Sub

Private Sub tmrCheckQueue_Timer()
    JPTRS.lblStatus.Caption = "Idle..."
    CheckQueue
End Sub
