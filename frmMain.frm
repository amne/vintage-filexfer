VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E6A71E71-C000-4629-8E97-05B3C3D464F6}#3.0#0"; "WinXock.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin WinXock.aWinXock WinXock 
      Height          =   555
      Left            =   0
      TabIndex        =   23
      Top             =   480
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   2040
      TabIndex        =   22
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Max             =   768
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   21
      Top             =   6435
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Text            =   "Server OFF"
            TextSave        =   "Server OFF"
            Object.Tag             =   "SrvOnOff"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Text            =   "Remote OFF"
            TextSave        =   "Remote OFF"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   4763
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "5:55 PM"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   0
      Top             =   1080
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect from Remote"
      Height          =   495
      Left            =   5760
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect to Remote"
      Height          =   495
      Left            =   5760
      TabIndex        =   19
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSendFile 
      Caption         =   "Send File"
      Height          =   255
      Left            =   6000
      TabIndex        =   18
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdStopServer 
      Caption         =   "Stop Server"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlgOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "FileXFer"
      Orientation     =   2
   End
   Begin VB.CommandButton cmdStartServer 
      Caption         =   "Start server"
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtChatReceive 
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3120
      Width           =   6855
   End
   Begin VB.TextBox txtRemotePort 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdSendChat 
      Caption         =   "Send"
      Height          =   735
      Left            =   6240
      TabIndex        =   13
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtRemoteIP 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtFLen 
      BackColor       =   &H80000011&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtFN 
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Top             =   360
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar pBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtChat 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5520
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   6120
   End
   Begin VB.Label lblSpeed 
      Caption         =   "Speed:"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblProgress 
      Caption         =   "Progress"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblRemotePort 
      Alignment       =   1  'Right Justify
      Caption         =   "Remote Port:"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblRemoteIP 
      Alignment       =   1  'Right Justify
      Caption         =   "Remote IP:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label lblFLen 
      Alignment       =   1  'Right Justify
      Caption         =   "File length:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblFn 
      Alignment       =   1  'Right Justify
      Caption         =   "File name:"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblChat 
      Caption         =   "Chat"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const lPort As Long = 51278
Private ReceivingFile As Boolean
Private SendingFile As Boolean
Private IAmServer As Boolean
Private recvFileID As Long
Private recvFName As String
Private sendFileID As Long
Private sendFName As String
Private sendBuffer As Long
Private bytesSend As Long
Private bytesToSend As Long
Private bytesRcvd As Long
Private bytesToRcv As Long

'Private Const OF_READ = &H0
'Private Const OF_READWRITE = &H2
'Private Const OF_SHARE_DENY_READ = &H30
'Private Const OF_SHARE_DENY_WRITE = &H20
'Private Const OF_WRITE = &H1
'Private Const FILE_BEGIN = 0
'Private Const FILE_CURRENT = 1
'Private Const FILE_END = 2
Private Const cstm_RCVD_FILE$ = "RCVD_FILE"
Private Const customSep$ = "+Êﬂ•+"


'Private Declare Function lwrite Lib "kernel32" Alias "_lwrite" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal wBytes As Long) As Long
'Private Declare Function lread Lib "kernel32" Alias "_lread" (ByVal hFile As Long, lpBuffer As Any, ByVal wBytes As Long) As Long
'Private Declare Function lopen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
'Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
'Private Declare Function llseek Lib "kernel32" Alias "_llseek" (ByVal hFile As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
'Private Declare Function lcreat Lib "kernel32" Alias "_lcreat" (ByVal lpPathName As String, ByVal iAttribute As Long) As Long


Public Sub WriteToFile(ByVal fID As Long, ByVal fData$, ByVal fLen&)
Put #fID, , fData
End Sub

Public Function ReadFromFile(ByVal fID As Long, ByVal fLen$) As String
Dim sTmp$
sTmp = String(fLen, " ")
Get #fID, , sTmp
ReadFromFile = sTmp
End Function

Private Sub cmdBrowse_Click()
cdlgOpen.ShowOpen
sendFName = cdlgOpen.FileName
sendFileID = FreeFile
Open sendFName For Binary Access Read Lock Read Write As sendFileID
txtFN.Text = sendFName
End Sub

Private Sub cmdConnect_Click()
Dim rConn$
rConn = InputBox("Enter the remote IP/hostname to connect to:" & vbCrLf & " (Ex: 12.34.56.78 or 127.0.0.1)", "FileXFer")
If rConn = "" Then Exit Sub
WinXock.xConnect rConn, lPort
End Sub

Private Sub cmdSendChat_Click()
xSendChat Trim(txtChat.Text)
End Sub

Private Sub cmdSendFile_Click()
Dim sBuf$
If sendFileID < 1 Then Exit Sub
bytesSend = 0
bytesRcvd = 0
bytesToSend = LOF(sendFileID)
sBuf = sendFName & customSep & bytesToSend & customSep
sendBuffer = Slider1.Value
If IAmServer Then
 WinXock.cl_XendData 2, 2, 1, 1, sBuf
 Else
 WinXock.cl_XendData 1, 2, 1, 1, sBuf
 End If
End Sub

Private Sub cmdStartServer_Click()
WinXock.StartServer lPort
WinXock.xConnect WinXock.GetLocalIP, lPort
IAmServer = True
WinXock.SetRefreshRate 50
End Sub

Private Sub cmdStopServer_Click()
Dim i&
For i = WinXock.GetSockLBound To WinXock.GetSockUBound
 WinXock.KillXock i
 Next i
WinXock.ShutDownServer True
IAmServer = False
WinXock.SetRefreshRate 0
End Sub

Private Sub Form_Load()
Slider1.Value = 128
End Sub

Private Sub Form_Unload(Cancel As Integer)
WinXock.ShutDownServer True
WinXock.ShutDownRefresh
WinXock.KillClient
WinXock.KillXock 1
WinXock.KillXock 2
End
End Sub

Private Sub Timer1_Timer()
Static bytesSent1&
Static bytesRcvd1&
On Local Error Resume Next
If SendingFile Then
 lblSpeed.Caption = "Speed: " & (bytesSend - bytesSent1) & " B/s"
 bytesSent1 = bytesSend
 pBar1.Value = (bytesSent1 / bytesToSend) * 100
 Else
 If ReceivingFile Then
  lblSpeed.Caption = "Speed: " & (bytesRcvd - bytesRcvd1) & " B/s"
  bytesRcvd1 = bytesRcvd
  pBar1.Value = (bytesRcvd1 / bytesToRcv) * 100
  Else
  lblSpeed.Caption = "Speed: 0 B/s"
  pBar1.Value = 0
  End If
 End If
End Sub

Private Sub Timer2_Timer()
If IAmServer Then
 StatusBar1.Panels(1).Text = "Server ON"
 StatusBar1.Panels(1).Enabled = True
 Else
 StatusBar1.Panels(1).Text = "Server OFF"
 StatusBar1.Panels(1).Enabled = False
 End If
cmdConnect.Enabled = (WinXock.GetClMember(sState) <> 7)
cmdDisconnect.Enabled = (WinXock.GetClMember(sState) <> 0)
cmdStartServer.Enabled = cmdConnect.Enabled
cmdStopServer.Enabled = (Not cmdStartServer.Enabled) And IAmServer

cmdSendFile.Enabled = Not (txtFN.Text = "")
If WinXock.GetXockMember(2, sState) = "7" Then
 StatusBar1.Panels(2).Text = "Remote ON"
 StatusBar1.Panels(2).Enabled = True
 Else
 StatusBar1.Panels(2).Text = "Remote OFF"
 StatusBar1.Panels(2).Enabled = False
 End If
End Sub

Public Sub Add2TextBox(tBox As TextBox, ByVal sTxt$, ByVal newLine As Boolean)
tBox.SelStart = Len(tBox.Text)
If newLine Then
 tBox.SelText = sTxt & vbCrLf
 Else
 tBox.SelLength = sTxt
 End If
tBox.SelStart = Len(tBox.Text)
End Sub

Private Sub ProcessRemote(xMsg As WinXock.TXMsg)
Dim ssPLs$(), nsPLs&
Dim sBuf$, bufLen&, msgTag&
ReDim fBuf(sendBuffer) As Byte
If xMsg.mTag = 1 Then
 Add2TextBox txtChatReceive, "[" & Time & "] <" & xMsg.mFrom & "> " & xMsg.mMsg, True
 Exit Sub
 End If
If xMsg.mTag = 2 Then
 ssPLs = Split(xMsg.mMsg, customSep)
 nsPLs = UBound(ssPLs)
 On Local Error Resume Next
 MsgBox "You are receiving a file from: " & txtRemoteIP.Text & vbCrLf & "You will now be prompted to save the file to a location on your computer!" & vbCrLf & "File name: " & ssPLs(0) & vbCrLf & "File size: " & ssPLs(1) & " bytes"
 On Local Error Resume Next
 cdlgOpen.FileName = ssPLs(0)
 cdlgOpen.ShowSave
 recvFName = cdlgOpen.FileName
 bytesToRcv = ssPLs(1)
 recvFileID = FreeFile
 Open recvFName For Binary Access Write Lock Read Write As recvFileID
 WinXock.cl_XendData xMsg.mFrom, 20, 1, 1, cstm_RCVD_FILE
 End If
If xMsg.mTag = 4 Then
 WriteToFile recvFileID, xMsg.mMsg, xMsg.mLen
 DoEvents
 WinXock.cl_XendData xMsg.mFrom, 20, 1, 1, cstm_RCVD_FILE
 ReceivingFile = True
 bytesRcvd = bytesRcvd + xMsg.mLen
 End If
If xMsg.mTag = 5 Then
 WriteToFile recvFileID, xMsg.mMsg, xMsg.mLen
 DoEvents
 bytesRcvd = bytesRcvd + xMsg.mLen
 Close recvFileID
 ReceivingFile = False
 MsgBox "File transfer complete!"
 End If
If xMsg.mTag = 20 Then
 SendingFile = True
 If xMsg.mMsg = cstm_RCVD_FILE Then
  sBuf = ReadFromFile(sendFileID, sendBuffer)
  msgTag = 4
  If Len(sBuf) > LOF(sendFileID) - bytesSend Then
   sBuf = Mid(sBuf, 1, LOF(sendFileID) - bytesSend)
   msgTag = 5
   Close sendFileID
   SendingFile = False
   End If
  DoEvents
  bufLen = Len(sBuf)
  WinXock.cl_XendData xMsg.mFrom, msgTag, 1, 1, sBuf
  bytesSend = bytesSend + bufLen
  End If
 End If
End Sub


Private Sub xSendChat(ByVal chatMessage$)
If Trim(txtChat.Text) <> "" Then
 If IAmServer Then
  WinXock.cl_XendData 2, 1, 1, 1, Trim(txtChat.Text)
  Else
  WinXock.cl_XendData 1, 1, 1, 1, Trim(txtChat.Text)
  End If
 End If
txtChat.SetFocus
txtChat.Text = ""
If IAmServer Then
 Add2TextBox txtChatReceive, "[" & Time & "] <" & 1 & "> " & chatMessage, True
 Else
 Add2TextBox txtChatReceive, "[" & Time & "] <" & 2 & "> " & chatMessage, True
 End If
End Sub

Private Sub txtChat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  xSendChat Trim(txtChat.Text)
  txtChat.SetFocus
  txtChat.Text = ""
 End If
End Sub

Private Sub WinXock_ConnxDataArrival(ByVal Index As Integer, xMsg As WinXock.TXMsg)
If Index = 0 Then
 ProcessRemote xMsg
 End If
End Sub

Private Sub WinXock_ConnxStateChange(ByVal Index As Integer, ByVal oldState As Long, ByVal newState As Long)
If Index = 2 And newState = 8 Then
 txtRemoteIP.Text = ""
 txtRemotePort.Text = ""
 WinXock.KillXock Index
 WinXock.StartServer lPort
 End If
End Sub

Private Sub WinXock_srvConnectionRequest(ByVal AssignedIndex As Long)
Dim rIP$, rPort$
rIP = WinXock.GetXockMember(AssignedIndex, sRemoteHostIP)
rPort = WinXock.GetXockMember(AssignedIndex, sRemotePort)
If AssignedIndex <> 1 Then
 txtRemoteIP.Text = rIP
 txtRemotePort.Text = rPort
 WinXock.ShutDownServer True
 MsgBox "Remote host connected! Ready to send file!"
 End If
End Sub
