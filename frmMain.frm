VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   FillColor       =   &H8000000F&
   ForeColor       =   &H80000004&
   LinkTopic       =   "Form1"
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      Height          =   615
      Left            =   5205
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
      Begin MSComctlLib.Toolbar tlbServerOptions 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "imlServerOptions"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
   End
   Begin VB.Frame fraCommand 
      Caption         =   "Command:"
      Height          =   615
      Left            =   45
      TabIndex        =   7
      Top             =   2880
      Width           =   5175
      Begin RichTextLib.RichTextBox rtbCommand 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   210
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   582
         _Version        =   393217
         BackColor       =   16777215
         TextRTF         =   $"frmMain.frx":0000
      End
   End
   Begin VB.Frame fraUsers 
      Caption         =   "People: 0"
      Height          =   2895
      Left            =   5205
      TabIndex        =   5
      Top             =   0
      Width           =   1695
      Begin MSComctlLib.TreeView trvBuddyList 
         Height          =   2295
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   4048
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imlBuddyList"
         BorderStyle     =   1
         Appearance      =   1
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
      Begin MSComctlLib.ProgressBar prbProgress 
         Height          =   180
         Left            =   40
         TabIndex        =   6
         Top             =   2640
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame fraServer 
      Caption         =   "Server Log"
      Height          =   2895
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   5175
      Begin RichTextLib.RichTextBox rtbServer 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4471
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":00BA
      End
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   3840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   4440
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlBuddyList 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0174
            Key             =   "Network"
            Object.Tag             =   "Network"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2928
            Key             =   "User"
            Object.Tag             =   "User"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlServerOptions 
      Left            =   2400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A84
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5238
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5554
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   rtbCommand.SelColor = vbBlack
   rtbCommand.SelFontName = "Small Fonts"
   rtbCommand.SelFontSize = 7
   ServerNick = "Server"
   MaxUsers = 5
   ReDim User.User(1 To MaxUsers)
   mCoolMenu.Install frmMenu.hwnd, , frmMenu.imlMenus
   mCoolMenu.SelectColor frmMenu.hwnd, QBColor(1)
   Set ToolTip = New frmTooltip
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   mCoolMenu.Uninstall frmMenu.hwnd
   ToolTip.DestroyTip
   End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mCoolMenu.Uninstall frmMenu.hwnd
   ToolTip.DestroyTip
   End
End Sub

Private Sub rtbCommand_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      
      Dim RichTextHolder As String
      Dim ServerCommand As String
      Dim ServerValue As String
      
      KeyAscii = 0
      
      If rtbCommand.Text = "" Then Exit Sub
         
      RichTextHolder = rtbCommand.Text
      ServerCommand = Trim(ParseString(RichTextHolder, "-", 1))
      ServerValue = Trim(ParseString(RichTextHolder, "-", 2))
      
      rtbCommand.Text = ""
      
      Select Case LCase(ServerCommand)
         Case "/start":
            
            Dim ServerIP As String
            Dim ServerPort As Integer
         
            If Len(ServerValue) > 4 Then
               AddSysMessage rtbServer, "*** Invalid Port ***", vbRed
               Exit Sub
            End If
            
            AddSysMessage rtbServer, "*** Server Started ***", vbBlack
            
            ServerPort = ServerValue
           
            sckListen.LocalPort = ServerPort
            sckListen.Close
            sckListen.Listen
            
            AddSysMessage rtbServer, "*** Listening on " & ServerValue & " ***", vbRed
         
         Case "/close":
            
            For SockInUse = 0 To MaxUsers
               sckServer(SockInUse).Close
            Next
         
            sckListen.Close
            
            SockInUse = 0
            
            AddSysMessage rtbServer, "*** Server Off ***", vbRed
         
         Case "/chat":
               
               For SeekIndex = 1 To MaxUsers
                  If User.User(SeekIndex).InChannel Then SocketSend SeekIndex, CHATJOIN & DELIMITER & ServerValue & DELIMITER & ServerNick
                  DoEvents
               Next
               
               ActiveChannel = ServerValue
               Me.WindowState = vbMinimized
               
               frmChat.Show
               
               frmChat.trvUsers.Nodes.Add , , , ServerNick, 1
               
               Dim InChannelCount As Integer
               
               For SeekIndex = 1 To MaxUsers
                  If User.User(SeekIndex).InChannel Then
                     InChannelCount = InChannelCount + 1
                  End If
               Next
               
               frmChat.fraUsers.Caption = "People: " & InChannelCount + 1
               
               For SeekIndex = 1 To MaxUsers
                  If User.User(SeekIndex).InChannel Then
                     frmChat.trvUsers.Nodes.Add , , , User.User(SeekIndex).UserName, 1
                  End If
               Next
               
         Case "/nick":
         
               ServerNick = ServerValue
               AddSysMessage rtbServer, "*** Nick changed to " & ServerNick & " ***", vbRed
      
      End Select
   End If
End Sub

Private Sub rtbCommand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   rtbCommand.SelColor = vbBlack
   rtbCommand.SelFontName = "Small Fonts"
   rtbCommand.SelFontSize = 7
End Sub

Private Sub rtbCommand_SelChange()
   rtbCommand.SelColor = vbBlack
   rtbCommand.SelFontName = "Small Fonts"
   rtbCommand.SelFontSize = 7
End Sub

Private Sub sckListen_Close()
   SockInUse = 0
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
   AcceptRequest requestID
End Sub

Private Sub sckServer_Close(Index As Integer)
   AddSysMessage rtbServer, "*** " & Trim(User.User(Index).UserName) & ": Closed Connection [ " & sckServer(Index).RemoteHostIP & " ] ***", vbRed
   
   If User.User(Index).InChannel Then
      For SeekIndex = 1 To frmChat.trvUsers.Nodes.Count
         If LCase(Trim(User.User(Index).UserName)) = LCase(Trim(frmChat.trvUsers.Nodes.Item(SeekIndex))) Then
            frmChat.trvUsers.Nodes.Remove (SeekIndex)
            frmChat.fraUsers.Caption = "People: " & frmChat.trvUsers.Nodes.Count
            AddSysMessage frmChat.rtbChat, "*** " & Trim(User.User(Index).UserName) & " left chat ***", vbRed
            Exit For
         End If
      Next
   
      For SeekIndex = 1 To MaxUsers
         If Not User.User(SeekIndex).UserIndex = Index And User.User(SeekIndex).InChannel Then SocketSend User.User(SeekIndex).UserIndex, CHATLEAVE & DELIMITER & User.User(SeekIndex).UserName
         DoEvents
      Next
   End If
      
      If frmMain.trvBuddyList.Nodes.Count = 0 Then Exit Sub
      
      For SeekIndex = 1 To trvBuddyList.Nodes.Count
         If LCase(Trim(trvBuddyList.Nodes.Item(SeekIndex).Text)) = LCase(Trim(User.User(Index).UserName)) Then
            trvBuddyList.Nodes.Remove (SeekIndex)
            User.UserCount = User.UserCount - 1
            fraUsers.Caption = "People: " & User.UserCount
            Exit For
         End If
      Next
            User.User(Index).Connected = False
            User.User(Index).InChannel = False
End Sub

Private Sub sckServer_Connect(Index As Integer)
   User.UserCount = User.UserCount + 1
   fraUsers.Caption = "People: " & User.UserCount
   SocketSend Index, SERVERMESSAGE & DELIMITER & "Welcome to my server " & sckServer(0).LocalIP
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
   Dim Message As String
   Dim ClientCommand As String

   sckServer(Index).GetData Message
   ClientCommand = Trim(ParseString(Message, DELIMITER, 1))
   
   Select Case ClientCommand
      Case USERINFORMATION:
      ' this is where you will determine pw vs sn
         User.User(Index).UserName = Trim(ParseString(Message, DELIMITER, 2))
         User.User(Index).info = ParseString(Message, DELIMITER, 3)
         User.User(Index).UserIndex = Index
         User.User(Index).Connected = True
         
         AddSysMessage rtbServer, "*** " & Trim(User.User(Index).UserName) & ": Connected [ " & sckServer(Index).RemoteHostIP & " ] ***", vbRed
         
         trvBuddyList.Nodes.Add , , , User.User(Index).UserName, 1
      
      Case IM:
      
      Case USERMESSAGE:
         Dim ToName As String
         Dim FromName As String
         Dim ImMessage As String
         Dim TheForm As Form
         
         ToName = Trim(ParseString(Message, DELIMITER, 2))
         FromName = Trim(ParseString(Message, DELIMITER, 3))
         ImMessage = Trim(ParseString(Message, DELIMITER, 4))

         If Not LCase(ToName) = LCase(ServerNick) Then
            For SeekIndex = 1 To MaxUsers
               If LCase(Trim(User.User(SeekIndex).UserName)) = LCase(Trim(ToName)) And User.User(SeekIndex).Connected Then
                  SendIm ToName, FromName, ImMessage, User.User(SeekIndex).UserIndex
                  Exit Sub
               End If
            Next
         End If
         
         For Each TheForm In Forms
            If TheForm.Caption = FromName Then
               Play "imrecieve.wav"
               AddChat TheForm.rtbIM, FromName, ImMessage, vbBlack
               Exit Sub
            End If
         Next
         
         Dim NewIm As New frmIM
         NewIm.Caption = FromName
         NewIm.Caption = FromName
         
         AddChat NewIm.rtbIM, FromName, ImMessage, vbBlack
         NewIm.Show
         
      Case CHATMESSAGE:
         Dim ChatNick As String
         Dim ChatNickMessage As String
         
         ChatNick = Trim(ParseString(Message, DELIMITER, 2))
         ChatNickMessage = Trim(ParseString(Message, DELIMITER, 3))
         
         If frmChat.Visible Then AddChat frmChat.rtbChat, ChatNick, ChatNickMessage, vbBlack
         
         For SeekIndex = 1 To MaxUsers
            If Not User.User(SeekIndex).UserIndex = Index And User.User(SeekIndex).Connected Then SocketSend SeekIndex, CHATMESSAGE & DELIMITER & ChatNick & DELIMITER & ChatNickMessage
            DoEvents
         Next
         
      Case PRIVATEMESSAGE:
         Dim PrivateString As String
         Dim PrivateToName As String
         Dim PrivateFromName As String
         
         PrivateToName = Trim(ParseString(Message, DELIMITER, 2))
         PrivateFromName = Trim(ParseString(Message, DELIMITER, 3))
         PrivateString = ParseString(Message, DELIMITER, 4)
         
         If Not PrivateToName = ServerNick Then
            For SeekIndex = 1 To MaxUsers
               If Trim(LCase(User.User(SeekIndex).UserName)) = Trim(LCase(PrivateToName)) And User.User(SeekIndex).InChannel Then SocketSend SeekIndex, PRIVATEMESSAGE & DELIMITER & PrivateFromName & DELIMITER & PrivateString
               DoEvents
            Next
         Else
            AddSysMessage frmChat.rtbChat, "*** Private Message [ " & PrivateFromName & " ] ***", vbRed
            AddChat frmChat.rtbChat, PrivateFromName, PrivateString, vbBlack
         End If
         
      Case CHATJOIN:
         Dim ChatJoinNick As String
         
         ActiveChannel = Trim(ParseString(Message, DELIMITER, 2))
         ChatJoinNick = Trim(ParseString(Message, DELIMITER, 3))
         
         If frmChat.Visible Then AddSysMessage frmChat.rtbChat, "*** " & ChatJoinNick & " entered chat ***", vbBlack
         If frmChat.Visible Then frmChat.trvUsers.Nodes.Add , , , ChatJoinNick, 1
         If frmChat.Visible Then frmChat.fraUsers.Caption = "People: " & frmChat.trvUsers.Nodes.Count
         
         User.User(Index).InChannel = True
         
         For SeekIndex = 1 To MaxUsers
            If Not User.User(SeekIndex).UserIndex = Index And User.User(SeekIndex).InChannel Then SocketSend SeekIndex, CHATJOIN & DELIMITER & ActiveChannel & DELIMITER & ChatJoinNick
            DoEvents
         Next
         
         Dim ChatUserNick As String
         
         For SeekIndex = 1 To MaxUsers
            If User.User(SeekIndex).InChannel Then
               If Not User.User(SeekIndex).UserIndex = Index Then
                  ChatUserNick = ChatUserNick & DELIMITER & User.User(SeekIndex).UserName
               End If
            End If
         Next
         
         If frmChat.Visible Then
            SocketSend Index, CHATUSERS & ChatUserNick & DELIMITER & ServerNick
         Else
            SocketSend Index, CHATUSERS & ChatUserNick
         End If
     
      Case CHATLEAVE:
         Dim ChatLeaveNick As String
         
         ChatLeaveNick = Trim(ParseString(Message, DELIMITER, 2))
         User.User(Index).InChannel = False
         
         If frmChat.Visible Then AddSysMessage frmChat.rtbChat, "*** " & ChatLeaveNick & " left chat ***", vbBlack
         
         For SeekIndex = 1 To frmChat.trvUsers.Nodes.Count
            If LCase(Trim(frmChat.trvUsers.Nodes.Item(SeekIndex).Text)) = LCase(Trim(ChatLeaveNick)) Then
               frmChat.trvUsers.Nodes.Remove (SeekIndex)
               Exit For
            End If
         Next
         
         If frmChat.Visible Then frmChat.fraUsers.Caption = "People: " & frmChat.trvUsers.Nodes.Count
         
         For SeekIndex = 1 To MaxUsers
            If Not User.User(SeekIndex).UserIndex = Index And User.User(SeekIndex).InChannel Then SocketSend SeekIndex, CHATLEAVE & DELIMITER & ChatLeaveNick
            DoEvents
         Next
         
      Case GETBUDDYSTATUS:
         Dim TheBuddy As String
         
         TheBuddy = Trim(ParseString(Message, DELIMITER, 2))
         
         For SeekIndex = 1 To trvBuddyList.Nodes.Count
            If LCase(TheBuddy) = LCase(Trim(trvBuddyList.Nodes.Item(SeekIndex))) Then
               SocketSend Index, BUDDYSTATUS & DELIMITER & ONLINE & DELIMITER & TheBuddy
               DoEvents
              Exit Sub
            End If
         Next
               SocketSend Index, BUDDYSTATUS & DELIMITER & OFFLINE & DELIMITER & TheBuddy
      
      Case KICKUSER:
         Dim KickNick As String
         Dim KickReason As String
         
         KickNick = Trim(ParseString(Message, DELIMITER, 2))
         KickReason = ParseString(Message, DELIMITER, 3)
         
         For SeekIndex = 1 To MaxUsers
            If LCase(Trim(User.User(SeekIndex).UserName)) = LCase(Trim(KickNick)) Then
               SocketSend User.User(SeekIndex).UserIndex, KICKUSER & DELIMITER & KickReason
               Exit Sub
            End If
         Next
   End Select
End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   AddSysMessage rtbServer, "*** " & User.User(Index).UserName & ": " & Description & " [ " & sckServer(Index).RemoteHostIP & " ] ***", vbRed
   
   If User.User(Index).InChannel Then
      For SeekIndex = 1 To frmChat.trvUsers.Nodes.Count
         If LCase(Trim(User.User(Index).UserName)) = LCase(Trim(frmChat.trvUsers.Nodes.Item(SeekIndex))) Then
            frmChat.trvUsers.Nodes.Remove (SeekIndex)
            frmChat.fraUsers.Caption = "People: " & frmChat.trvUsers.Nodes.Count
            AddSysMessage frmChat.rtbChat, "*** " & Trim(User.User(Index).UserName) & " left chat ***", vbRed
            Exit For
         End If
      Next
   
      For SeekIndex = 1 To MaxUsers
         If Not User.User(SeekIndex).UserIndex = Index And User.User(SeekIndex).InChannel Then SocketSend SeekIndex, CHATLEAVE & DELIMITER & User.User(SeekIndex).UserName
         DoEvents
      Next
   End If
      
      For SeekIndex = 1 To MaxUsers
         If LCase(Trim(trvBuddyList.Nodes.Item(SeekIndex).Text)) = LCase(Trim(User.User(Index).UserName)) Then
            trvBuddyList.Nodes.Remove (SeekIndex)
            User.UserCount = User.UserCount - 1
            fraUsers.Caption = "People: " & User.UserCount
            Exit For
         End If
      Next
            User.User(Index).Connected = False
            User.User(Index).InChannel = False
End Sub

Private Sub trvBuddyList_DblClick()
   If trvBuddyList.SelectedItem Is Nothing Then Exit Sub
   Dim NewIm As New frmIM
   
   NewIm.Caption = Trim(trvBuddyList.SelectedItem.Text)
   NewIm.Show
End Sub

Private Sub trvBuddyList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      Me.PopupMenu frmMenu.mnuList
   Else
      If trvBuddyList.SelectedItem Is Nothing Then Exit Sub
      'does not show tip for user after sign off and users are left
      For SeekIndex = 1 To MaxUsers
         If LCase(Trim(User.User(SeekIndex).UserName)) = LCase(Trim(trvBuddyList.SelectedItem.Text)) Then
            ToolTip.HideTip
            ToolTip.Text = "Nick: " & User.User(SeekIndex).UserName & vbCrLf & "IP: " & sckServer(User.User(SeekIndex).UserIndex).RemoteHostIP & vbCrLf & "Level: " & User.User(SeekIndex).UserIndex
            ToolTip.ToolBackColor = QBColor(1)
            ToolTip.ToolFont = "Arial"
            ToolTip.ToolFontSize = 10
            'ToolTip.ToolFontBold = True
            ToolTip.ToolForeColor = vbWhite
            ToolTip.ShowTime = 10000
            ToolTip.CreateTip trvBuddyList
            ToolTip.ShowTip
         End If
      Next
   End If
End Sub

Public Sub AcceptRequest(SocketID As Long)
   Dim SockIndex As Integer
   
   For SockIndex = 1 To SockInUse
      If Not sckServer(SockIndex).State = sckConnected Then
         sckServer(SockIndex).Close
         sckServer(SockIndex).Accept SocketID
         sckServer_Connect (SockIndex)
         Exit Sub
      End If
   Next

   If SockInUse = MaxUsers Then Exit Sub

   SockInUse = SockInUse + 1
   Load frmMain.sckServer(SockInUse)
   sckServer(SockInUse).Close
   sckServer(SockInUse).Accept SocketID
   sckServer_Connect (SockInUse)
End Sub
