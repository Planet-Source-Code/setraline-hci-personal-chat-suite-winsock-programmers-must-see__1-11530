VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Main"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   Picture         =   "frmMainClient.frx":0000
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prbProgress 
      Height          =   180
      Left            =   420
      TabIndex        =   5
      Top             =   3900
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "x"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "check out the chat part that works now"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   5040
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlBuddyList 
      Left            =   4920
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbCommand 
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   3420
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      MultiLine       =   0   'False
      MaxLength       =   50
      Appearance      =   0
      TextRTF         =   $"frmMainClient.frx":6719E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbServer 
      Height          =   2175
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMainClient.frx":6726D
   End
   Begin MSComctlLib.TreeView trvBuddyList 
      Height          =   2295
      Left            =   5760
      TabIndex        =   0
      Top             =   405
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   4048
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
   frmChat.Show
End Sub

Private Sub Command2_Click()
   End
End Sub

Private Sub Form_Load()
   'AutoFormShape Me, vbWhite
   SetTVBackColour trvBuddyList, vbBlack
   rtbCommand.SelBold = False
   rtbCommand.SelUnderline = False
   rtbCommand.SelItalic = False
   rtbCommand.SelColor = vbYellow
   rtbCommand.SelFontName = "Verdana"
   rtbCommand.SelFontSize = 7
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim ReturnVal As Long
   
   ReleaseCapture
   ReturnVal = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub rtbCommand_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      
      Dim ClientComs As String
      Dim ClientCommand As String
      Dim ClientValue As String
      
      KeyAscii = 0
      
      If Not rtbCommand.Text <> "" Then Exit Sub
      
      ClientComs = rtbCommand.Text
      ClientCommand = ParseString(ClientComs, "#", 1)
      ClientValue = ParseString(ClientComs, "#", 2)
      
      rtbCommand.Text = ""
      
      Select Case LCase(ClientCommand)
            Case LCase("/connect "):
               
               Dim ServerIP As String
               Dim ServerPort As Integer
         
               AddSysMessage rtbServer, "*** Connecting to " & ClientValue & " ***"
               ServerIP = ClientValue
               ServerPort = 3370
               sckClient.Close
               sckClient.Connect ServerIP, ServerPort
               
            Case LCase("/msg "):
               sckClient.SendData USERMESSAGE & DELIMITER & ClientValue
               Unload frmIM
               frmIM.lblIm.Caption = "Instant Message To: " & ClientValue
               frmIM.Show
               
            Case LCase("/chat "):
               sckClient.SendData CHAT & DELIMITER & ClientValue
               Unload frmChat
               ActiveChannel = ClientValue
               frmChat.lblChat.Caption = "Channel - " & ActiveChannel
               frmChat.Show
               Me.WindowState = vbMinimized
     End Select
   End If
End Sub

Private Sub rtbCommand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   rtbCommand.SelBold = False
   rtbCommand.SelUnderline = False
   rtbCommand.SelItalic = False
   rtbCommand.SelColor = vbYellow
   rtbCommand.SelFontName = "Verdana"
   rtbCommand.SelFontSize = 7
End Sub

Private Sub rtbCommand_SelChange()
   rtbCommand.SelBold = False
   rtbCommand.SelUnderline = False
   rtbCommand.SelItalic = False
   rtbCommand.SelColor = vbYellow
   rtbCommand.SelFontName = "Verdana"
   rtbCommand.SelFontSize = 7
End Sub

Private Sub sckClient_Close()
   AddSysMessage rtbServer, "*** Connection Closed ***"
End Sub

Private Sub sckClient_Connect()
   AddSysMessage rtbServer, "*** Connected to Server ***"
   sckClient.SendData USERINFORMATION & DELIMITER & "Setraline" & DELIMITER & "Im the man you know it"
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
   Dim Message As String
   Dim RemoteCommand As String
   
   sckClient.GetData Message
   RemoteCommand = ParseString(Message, DELIMITER, 1)
   
   Select Case RemoteCommand
      Case SERVERMESSAGE:
         AddSysMessage rtbServer, "*** " & ParseString(Message, DELIMITER, 2) & " ***", vbGreen
      
      Case CHAT:
         Unload frmChat
         ActiveChannel = ParseString(Message, DELIMITER, 2)
         frmChat.lblChat.Caption = "Channel - " & ChannelName
         frmChat.Show
         Me.WindowState = vbMinimized
      
      Case CHATJOIN:
         
      Case CHATLEAVE:
      
      Case USERMESSAGE:
         Unload frmIM
         frmIM.lblIm.Caption = "Instant Message From: " & ParseString(Message, DELIMITER, 2)
         frmIM.Show
      
      Case CHATMESSAGE:
         AddChat frmChat.rtbChat, ParseString(Message, DELIMITER, 2), ParseString(Message, DELIMITER, 3)
      
      Case NEWUSER:
      
      Case KICKUSER:
      
      Case NICKINUSE:
   End Select
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   AddSysMessage rtbServer, "*** Error Connecting ***"
End Sub


