VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtX 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   2760
      Width           =   10095
   End
   Begin VB.Frame framePorts 
      Caption         =   "Settings"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin MSWinsockLib.Winsock tcpExtra 
         Left            =   2760
         Top             =   1800
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtRemoteHost 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Text            =   "192.168.0.74"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Listen"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtLocalPort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "3389"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtRemotePort 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "40001"
         Top             =   945
         Width           =   1455
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Client"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Server"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin MSWinsockLib.Winsock tcpLocal 
         Left            =   2760
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock tcpRemote 
         Left            =   2760
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblRemoteHost 
         Alignment       =   1  'Right Justify
         Caption         =   "Remote Host:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblLocal 
         Caption         =   "Waiting..."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label lblRemote 
         BackStyle       =   0  'Transparent
         Caption         =   "Waiting..."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label lblRemoteStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Remote Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblLocalStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblLocalPort 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label lblRemotePort 
         Alignment       =   1  'Right Justify
         Caption         =   "Remote Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()
    optMode(0).Enabled = False
    optMode(1).Enabled = False
    cmdConnect.Enabled = False
    txtRemotePort.Enabled = False
    txtLocalPort.Enabled = False
    txtRemoteHost.Enabled = False
    
    ' Connect to the client
    tcpRemote.Connect txtRemoteHost.Text, CLng(txtRemotePort.Text)
    UpdateRemote "Connecting...", vbCyan
End Sub

Private Sub cmdListen_Click()
    optMode(0).Enabled = False
    optMode(1).Enabled = False
    cmdListen.Enabled = False
    txtRemotePort.Enabled = False
    txtLocalPort.Enabled = False
    
    ' Prepare the socket to listen remotely on an open port.
    tcpRemote.Bind CLng(txtRemotePort.Text)
    tcpRemote.Listen
    
    UpdateRemote "Listening on " & txtRemotePort.Text & "...", vbBlue
End Sub

Private Sub optMode_Click(Index As Integer)
    If optMode(0).Value = True Then
        cmdListen.Enabled = False
        cmdConnect.Enabled = True
        txtRemoteHost.Enabled = True
    Else
        cmdListen.Enabled = True
        cmdConnect.Enabled = False
        txtRemoteHost.Enabled = False
    End If
End Sub

Private Sub tcpLocal_Close()
    UpdateLocal "Socket closed", vbRed
End Sub

Private Sub tcpLocal_Connect()
    UpdateLocal "Connected [" & tcpLocal.RemoteHostIP & ":" & tcpLocal.RemotePort & "]", vbBlue
End Sub

Private Sub tcpLocal_ConnectionRequest(ByVal requestID As Long)
    ' Accept the connection from the local program.
    tcpLocal.Close
    tcpLocal.Accept requestID
    UpdateLocal "Connected [" & tcpLocal.RemoteHostIP & ":" & tcpLocal.RemotePort & "]", vbBlue
End Sub

Private Sub tcpLocal_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    ' Accept the incoming data from the program.
    tcpLocal.GetData strData, vbString
    AddText "- Got " & Len(strData) & " bytes locally."
    AddText DebugOutput(strData)
    
    ' Send the data out to the server
    tcpRemote.SendData strData
    AddText "- Sent " & Len(strData) & " bytes remotely."
End Sub

Private Sub tcpLocal_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    tcpLocal.Close
    UpdateLocal "Socket error", vbRed
    AddText "[Local] Socket error (" & Number & "): " & Description
End Sub

Private Sub tcpRemote_Close()
    UpdateRemote "Socket closed", vbRed
End Sub

Private Sub tcpRemote_Connect()
    ' We connected to the client.
    UpdateRemote "Connected [" & tcpRemote.RemoteHostIP & ":" & tcpRemote.RemotePort & "]", vbBlue
    
    ' Initiate connection to the server process.
    tcpLocal.Connect "localhost", CLng(txtLocalPort.Text)
    UpdateLocal "Connecting...", vbCyan
End Sub

Private Sub tcpRemote_ConnectionRequest(ByVal requestID As Long)
    ' Accept the connection from the server
    tcpRemote.Close
    tcpRemote.Accept requestID
    UpdateRemote "Connected [" & tcpRemote.RemoteHostIP & ":" & tcpRemote.RemotePort & "]", vbBlue
    
    ' Prepare the local socket to receive the connection.
    tcpLocal.Bind CLng(txtLocalPort.Text)
    tcpLocal.Listen
    
    ' At this point, we should connect our local program.
    UpdateLocal "Listening on " & txtLocalPort.Text & "...", vbBlue
End Sub

Private Sub UpdateLocal(ByVal Status As String, ByVal Color As Long)
    lblLocal.Caption = Status
    lblLocal.ForeColor = Color
End Sub

Private Sub UpdateRemote(ByVal Status As String, ByVal Color As Long)
    lblRemote.Caption = Status
    lblRemote.ForeColor = Color
End Sub

Private Sub tcpRemote_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    ' Accept the incoming data from the other computer
    tcpRemote.GetData strData, vbString
    AddText "-- Got " & Len(strData) & " bytes remotely."
    AddText DebugOutput(strData)
    
    ' Send the data out to the program
    tcpLocal.SendData strData
    AddText "-- Sent " & Len(strData) & " bytes locally."
End Sub

Private Sub tcpRemote_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    tcpRemote.Close
    UpdateRemote "Socket error", vbRed
    AddText "[Remote] Socket error (" & Number & "): " & Description
End Sub

Private Sub AddText(ByVal Text As String)
    txtX.Text = txtX.Text & Text & vbCrLf
    txtX.SelStart = Len(txtX.Text)
End Sub

Public Function DebugOutput(ByVal sIn As String) As String
   Dim x1 As Long, y1 As Long
   Dim iLen As Long, iPos As Long
   Dim sB As String, sT As String
   Dim sOut As String
   Dim Offset As Long, sOffset As String
   
   iLen = Len(sIn)
   If iLen = 0 Then Exit Function
   sOut = ""
   Offset = 0
   For x1 = 0 To ((iLen - 1) \ 16)
       sOffset = Right$("0000" & Hex(Offset), 4)
       sB = String(48, " ")
       sT = "................"
       For y1 = 1 To 16
           iPos = 16 * x1 + y1
           If iPos > iLen Then Exit For
           Mid(sB, 3 * (y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
           Select Case Asc(Mid(sIn, iPos, 1))
           Case 0, 9, 10, 13
           Case Else
               Mid(sT, y1, 1) = Mid(sIn, iPos, 1)
           End Select
       Next y1
       If Len(sOut) > 0 Then sOut = sOut & vbCrLf
       sOut = sOut & sOffset & ":  "
       sOut = sOut & sB & "  " & sT
       Offset = Offset + 16
   Next x1
   DebugOutput = sOut
End Function
