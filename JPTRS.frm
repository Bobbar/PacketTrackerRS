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


Private Sub chkVerbose_Click()
bolVerbose = CBool(chkVerbose.Value)
End Sub

Private Sub cmdCommand1_Click()
Debug.Print DateTime.Date & " " & DateTime.Time

End Sub

Private Sub Form_Load()
bolVerbose = CBool(chkVerbose.Value)
strLogLoc = Environ$("APPDATA") & "\JPTRS\LOG.LOG"

    FindMySQLDriver
cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"
GetUserIndex
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

cn_global.Close
Unload Me
End

End Sub

Private Sub tmrCheckQueue_Timer()
JPTRS.lblStatus.Caption = "Idle..."
CheckQueue





End Sub
