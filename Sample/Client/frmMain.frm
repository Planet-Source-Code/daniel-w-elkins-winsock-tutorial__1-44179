VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daytime Client"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Left            =   5400
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status : Idle."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   6495
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   60
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Port :"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Server :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
If Len(txtServer.Text) = 0 Then
MsgBox "Enter a server to connect to", vbCritical, "Server Required"
txtServer.SetFocus
Exit Sub
ElseIf Len(txtPort.Text) = 0 Then
MsgBox "Enter a port to connect on", vbCritical, "Port Required"
txtPort.SetFocus
Exit Sub
ElseIf Not IsNumeric(txtPort.Text) Then
MsgBox "Enter a numeric value for the port", vbCritical, "Invalid Port Value"
txtPort.SelStart = 0
txtPort.SelLength = Len(txtPort.Text)
Exit Sub
End If
Socket.Close
Socket.Connect txtServer.Text, txtPort.Text
lblStatus.Caption = "Status : Connecting . . ."
End Sub

Private Sub Socket_Close()
lblStatus.Caption = "Status : Session Complete."
End Sub

Private Sub Socket_Connect()
lblStatus.Caption = "Status : Retrieving Date / Time . . ."
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim sDat As String: sDat = Empty
Dim sBuff() As String
Socket.GetData sDat
sBuff() = Split(sDat, "***")
lblDate.Caption = "Server Date : " & sBuff(0)
lblTime.Caption = "Server Time : " & sBuff(1)
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
lblStatus.Caption = "Status : Error Connecting (" & Description & " - #" & Number & ")."
End Sub
