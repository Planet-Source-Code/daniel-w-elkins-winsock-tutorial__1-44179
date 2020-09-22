VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daytime Server"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4485
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
   ScaleHeight     =   3960
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   840
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkStop 
      Caption         =   "S&top Server"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "&Start Server"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin MSComctlLib.ListView LVCon 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   " Right Click for More Options "
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IP Address"
         Object.Width           =   3122
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date / Time"
         Object.Width           =   4154
      EndProperty
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   4680
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1234
   End
   Begin VB.Label lblHits 
      Caption         =   "0"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Connections (Hits) :"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connection Log :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mnuCon 
      Caption         =   "Con"
      Visible         =   0   'False
      Begin VB.Menu mnuClearList 
         Caption         =   "Clear List"
      End
      Begin VB.Menu mnuSaveList 
         Caption         =   "Save List"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+                                                                                                                                                              +
'+                  All of the lines commented in green are the examples from the tutorial                              +
'+                                                                                                                                                              +
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


Option Explicit

Dim lHits As Long
Dim lCurItem As Long

Private Sub chkStart_Click()
If chkStart.Value Then
Socket.Close
Socket.Listen
chkStart.Enabled = False
chkStop.Value = 1
chkStop.Enabled = True
chkStop.Value = 0
End If

End Sub

Private Sub chkStop_Click()
If chkStop.Value Then
Socket.Close
chkStop.Enabled = False
chkStart.Enabled = True
chkStop.Value = 1
chkStart.Value = 0
End If
End Sub

Private Sub cmdReset_Click()
lHits = 0
lblHits.Caption = "Connections (Hits) : 0"
End Sub

Private Sub Form_Load()
Socket.Listen ''///
End Sub

Private Sub LVCon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then Me.PopupMenu mnuCon
End Sub

Private Sub mnuClearList_Click()
Dim Rep As Integer
Rep = MsgBox("Are you sure you want to clear the list ?", vbQuestion + vbYesNo, "Clear Connection Log")
If Rep = 6 Then LVCon.ListItems.Clear: lCurItem = 0
End Sub

Private Sub mnuSaveList_Click()
Dim FF As Integer: FF = FreeFile
Dim lItem As Long: lItem = 0
Dim sTmpDat As String: sTmpDat = Empty
With CD
.DialogTitle = "Save Connection Log"
.Filter = "All Supported File Types|*.txt;*.dat;*.tmp;*.log;*.usr|Text Files|*.txt|Data Files|*.dat|Temporary Files|*.tmp|Log Files|*.log|User Files|*.usr"
.ShowSave
If Len(.FileName) > 0 Then
For lItem = 1 To LVCon.ListItems.Count
sTmpDat = sTmpDat & "IP Address : " & LVCon.ListItems(lItem).Text & vbNewLine & "Date / Time : " & LVCon.ListItems(lItem).ListSubItems(1).Text & vbNewLine & vbNewLine
DoEvents
Next lItem
sTmpDat = Mid(sTmpDat, 1, Len(sTmpDat) - 2)
Open .FileName For Binary As #1
Close #1
Open .FileName For Binary Access Write As #FF
Put #FF, , sTmpDat
Close #FF
End If
End With
sTmpDat = Empty
End Sub

Private Sub Socket_Close()
'(This isn't in the tutorial)
'as soon as the client disconnects, the server 're-opens' so another client can connect
Socket.Close
Socket.Listen

End Sub

Private Sub Socket_ConnectionRequest(ByVal requestID As Long)
Socket.Close '///
Socket.Accept requestID
lHits = lHits + 1
lblHits.Caption = lHits
LVCon.ListItems.Add , , Socket.RemoteHostIP
lCurItem = lCurItem + 1
LVCon.ListItems(lCurItem).ListSubItems.Add , , Now
Socket.SendData Date & "***" & Time '///
End Sub

Private Sub Socket_SendComplete()
Socket.Close '///
Socket.Listen
End Sub
