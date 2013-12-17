VERSION 5.00
Begin VB.Form JPTRS 
   AutoRedraw      =   -1  'True
   Caption         =   "JPTRS"
   ClientHeight    =   5235
   ClientLeft      =   120
   ClientTop       =   450
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
   ScaleHeight     =   5235
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   3240
      Picture         =   "JPTRS.frx":08CA
      Top             =   1020
      Width           =   1350
   End
End
Attribute VB_Name = "JPTRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    FindMySQLDriver
cn_global.Open "uid=" & strUsername & ";pwd=" & strPassword & ";server=" & strServerAddress & ";" & "driver={" & strSQLDriver & "};database=TicketDB;dsn=;"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

cn_global.Close
Unload Me
End

End Sub

