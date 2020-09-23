VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBuddy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buddy List"
   ClientHeight    =   3015
   ClientLeft      =   2220
   ClientTop       =   1050
   ClientWidth     =   2070
   FillColor       =   &H8000000F&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   138
   Begin VB.Frame Frame1 
      Height          =   180
      Left            =   0
      TabIndex        =   5
      Top             =   -75
      Width           =   2055
   End
   Begin VB.Timer tmrBuddyStatus 
      Interval        =   5000
      Left            =   1320
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock sckSendFile 
      Left            =   0
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdgFileDialog 
      Left            =   360
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   840
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabBuddy 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Online"
      TabPicture(0)   =   "frmBuddy.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "prbProgress"
      Tab(0).Control(1)=   "trvBuddies"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "List Setup"
      TabPicture(1)   =   "frmBuddy.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "trvEditBuddies"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAddGroup"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdAddBuddy"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdDelete"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H8000000C&
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdAddBuddy 
         BackColor       =   &H8000000C&
         Caption         =   "Add B"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdAddGroup 
         BackColor       =   &H8000000C&
         Caption         =   "Add G"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "+group"
         Height          =   375
         Left            =   -74160
         TabIndex        =   3
         Top             =   3120
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+buddy"
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   3120
         Width           =   735
      End
      Begin MSComctlLib.TreeView trvBuddies 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   0
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3413
         _Version        =   393217
         Indentation     =   88
         LabelEdit       =   1
         Style           =   5
         ImageList       =   "imgBuddy"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView trvEditBuddies 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   3201
         _Version        =   393217
         Indentation     =   88
         LabelEdit       =   1
         Style           =   5
         ImageList       =   "imgBuddy"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ProgressBar prbProgress 
         Height          =   180
         Left            =   -74880
         TabIndex        =   6
         Top             =   2520
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.ImageList imgBuddy 
      Left            =   1680
      Top             =   3840
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
            Picture         =   "frmBuddy.frx":0038
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":0192
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBuddy.frx":02EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
      End
   End
End
Attribute VB_Name = "frmBuddy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddBuddy_Click()
  Dim nodBuddy As Node
  If ExistsInTree(trvEditBuddies, "New Buddy", True) = False Then
    If trvEditBuddies.Nodes.Count < 1 Then
      MsgBox "You need a group to add buddies to.", vbOKOnly + vbCritical, "Error": Exit Sub
      Exit Sub
    End If
    If trvEditBuddies.SelectedItem Is Nothing Then
      Set nodBuddy = trvEditBuddies.Nodes.Add(1, tvwChild, , "New Buddy", 3, 3)
    Else
      If trvEditBuddies.SelectedItem.Parent Is Nothing Then
        Set nodBuddy = trvEditBuddies.Nodes.Add(trvEditBuddies.SelectedItem.Index, tvwChild, , "New Buddy", 3, 3)
      Else
        Set nodBuddy = trvEditBuddies.Nodes.Add(trvEditBuddies.SelectedItem.Index, tvwPrevious, , "New Buddy", 3, 3)
      End If
    End If
    nodBuddy.Selected = True
    trvEditBuddies.SetFocus
    trvEditBuddies.StartLabelEdit
  End If
End Sub

Private Sub cmdAddGroup_Click()
  Dim lngCounter As Long, strKey As String, nodGroup As Node
  If ExistsInTree(trvEditBuddies, "New Group", True) = False Then
    If trvEditBuddies.SelectedItem Is Nothing Then
      Set nodGroup = trvEditBuddies.Nodes.Add(, , , "Forem", 1, 1)
    Else
      If trvEditBuddies.SelectedItem.Parent Is Nothing Then
        Set nodGroup = trvEditBuddies.Nodes.Add(trvEditBuddies.SelectedItem.Index, tvwNext, , "Forem", 1, 1)
      Else
        Set nodGroup = trvEditBuddies.Nodes.Add(trvEditBuddies.SelectedItem.Parent.Index, tvwNext, , "Forem", 1, 1)
      End If
    End If
      nodGroup.Selected = True
      trvEditBuddies.SetFocus
    'trvEditBuddies.StartLabelEdit
    'trvEditBuddies.SelectedItem.Bold = True
    'set this as default for now
      nodGroup.Bold = True
      nodGroup.Expanded = True
      cmdAddGroup.Enabled = False ' just for now
  End If
End Sub

Private Sub cmdDelete_Click()
   Dim TreeItem As Integer
   Dim BuddyIndex As Integer
   Dim BuddyName As String

   TreeItem = trvEditBuddies.SelectedItem.Index
   BuddyName = Trim(trvEditBuddies.Nodes.Item(TreeItem).Text)
   trvEditBuddies.Nodes.Remove (TreeItem)

   For BuddyIndex = 1 To 10
      If LCase(Trim(Buddy.Buddy(BuddyIndex).BuddyNick)) = LCase(BuddyName) Then
         Buddy.Buddy(BuddyIndex).BuddyNick = ""
         Exit For
      End If
   Next
   
   For BuddyIndex = 1 To frmBuddy.trvBuddies.Nodes.Count
      If LCase(Trim(frmBuddy.trvBuddies.Nodes.Item(BuddyIndex).Text)) = LCase(BuddyName) Then
         frmBuddy.trvBuddies.Nodes.Remove (BuddyIndex)
      End If
   Next
   
   SaveBuddies
End Sub

Private Sub Form_Load()
   mCoolMenu.Install frmMenu.hWnd, , frmMenu.imlMenus
   mCoolMenu.SelectColor frmMenu.hWnd, QBColor(1)
   LoadBuddies
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mCoolMenu.Uninstall frmMenu.hWnd
   End
End Sub

Private Sub mnuSettings_Click()
   MsgBox "Not done yet", vbInformation, "Not done yet"
End Sub

Private Sub sckClient_Close()
   Play "disconnect.wav"
   frmChat.trvUsers.Nodes.Clear
   frmChat.fraUsers.Caption = "People: 0"
   For SeekIndex = 1 To trvBuddies.Nodes.Count
      If Not trvBuddies.Nodes.Item(SeekIndex).Bold Then
         trvBuddies.Nodes.Remove (SeekIndex)
      End If
   Next
   Me.Hide
   frmSignOn.Show
End Sub

Private Sub sckClient_Connect()
   SocketSend USERINFORMATION & DELIMITER & ClientNick & DELIMITER & "Forem 2000 Client v1"
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
   Dim Message As String
   Dim RemoteCommand As String
   
   sckClient.GetData Message, vbString
   RemoteCommand = Trim(ParseString(Message, DELIMITER, 1))
   
   Select Case RemoteCommand
      Case SERVERMESSAGE:
         Dim ServerConnectMessage As String
         
         ServerConnectMessage = ParseString(Message, DELIMITER, 2)
         
      Case IM:
    
      Case CHATJOIN:
         Dim ChatJoinNick As String
         
         ActiveChannel = Trim(ParseString(Message, DELIMITER, 2))
         ChatJoinNick = Trim(ParseString(Message, DELIMITER, 3))
         
         If Not frmChat.Visible Then frmChat.Show
         If Not Me.WindowState = vbMinimized Then Me.WindowState = vbMinimized
         
         frmChat.trvUsers.Nodes.Add , , , ChatJoinNick, 1
         AddSysMessage frmChat.rtbChat, "*** " & ChatJoinNick & " entered chat ***", vbBlack
         frmChat.fraUsers.Caption = "People: " & frmChat.trvUsers.Nodes.Count
         Play "chatjoin.wav"
         
      Case CHATLEAVE:
         Dim ChatLeaveNick As String
         
         ChatLeaveNick = ParseString(Message, DELIMITER, 2)
         
         For SeekIndex = 1 To frmChat.trvUsers.Nodes.Count
            If frmChat.trvUsers.Nodes.Item(SeekIndex).Text = ChatLeaveNick Then
               frmChat.trvUsers.Nodes.Remove (SeekIndex)
               Exit For
            End If
         Next
         
         AddSysMessage frmChat.rtbChat, "*** " & ChatLeaveNick & " left chat ***", vbRed
         frmChat.fraUsers.Caption = "People: " & frmChat.trvUsers.Nodes.Count
         Play "chatleave.wav"
         
      Case CHATUSERS:
         Dim ChatNickAdd As String
         Dim DelimiterCount As Integer
         
         DelimiterCount = 2
         
         Do
            ChatNickAdd = Trim(ParseString(Message, DELIMITER, DelimiterCount))
            DelimiterCount = DelimiterCount + 1
            If Not ChatNickAdd = "" Then frmChat.trvUsers.Nodes.Add , , , ChatNickAdd, 1
         Loop Until ChatNickAdd = ""
         
         frmChat.fraUsers.Caption = "People: " & frmChat.trvUsers.Nodes.Count
         
      Case USERMESSAGE:
         Dim ToName As String
         Dim FromName As String
         Dim ImMessage As String
         Dim TheForm As Form
         
         ToName = Trim(ParseString(Message, DELIMITER, 2))
         FromName = Trim(ParseString(Message, DELIMITER, 3))
         ImMessage = ParseString(Message, DELIMITER, 4)

         For Each TheForm In Forms
            If TheForm.Caption = FromName Then
               AddChat TheForm.rtbIM, FromName, ImMessage, vbBlack
               Play "imrecieve.wav"
               Exit Sub
            End If
         Next
         
         Dim NewIm As New frmIM
         NewIm.Caption = FromName
         AddChat NewIm.rtbIM, FromName, ImMessage, vbBlack
         Play "imrecieve.wav"
         NewIm.Show
      
      Case CHATMESSAGE:
         Dim ChatString As String
         Dim ChatName As String
         
         ChatName = Trim(ParseString(Message, DELIMITER, 2))
         ChatString = ParseString(Message, DELIMITER, 3)
         AddChat frmChat.rtbChat, ChatName, ChatString, vbBlack
         
      Case PRIVATEMESSAGE:
         Dim PrivateString As String
         Dim PrivateName As String
         
         PrivateName = Trim(ParseString(Message, DELIMITER, 2))
         PrivateString = ParseString(Message, DELIMITER, 3)
         AddSysMessage frmChat.rtbChat, "*** Private Message [ " & PrivateName & " ] ***", vbRed
         AddChat frmChat.rtbChat, PrivateName, PrivateString, vbBlack
         Play "privatemessage.wav"
      
      Case NEWUSER:
      
      Case KICKUSER:
         Dim KickReason As String
         
         KickReason = Trim(ParseString(Message, DELIMITER, 2))
         Unload frmChat
         MsgBox "You have been kicked from chat [ " & KickReason & " ]", vbInformation, "Kicked"
      
      Case NICKINUSE:
      
      Case BUDDYSTATUS:
         Dim UserStatus As String
         Dim BuddyName As String
         Dim BuddyStringError As String
         
         UserStatus = Trim(ParseString(Message, DELIMITER, 2))
         BuddyName = Trim(ParseString(Message, DELIMITER, 3))
         BuddyStringError = Trim(ParseString(Message, DELIMITER, 4))
         
         'If Not BuddyStringError = "" Then Exit Sub
         
         If UserStatus = ONLINE Then
            If ExistsInTree(trvBuddies, BuddyName, False, False) Then Exit Sub
            trvBuddies.Nodes.Add nodGroup, tvwChild, , BuddyName, 3
            Play "buddyin.wav"
         Else
            If ExistsInTree(trvBuddies, BuddyName, False, True) Then Play "buddyout.wav"
         End If
   End Select
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   Play "disconnect.wav"
   MsgBox Description, vbCritical, "Error"
   mCoolMenu.Uninstall frmMenu.hWnd
   frmChat.trvUsers.Nodes.Clear
   frmChat.fraUsers.Caption = "People: 0"
   Me.Hide
   frmSignOn.Show
End Sub

Private Sub sckClient_SendComplete()
   SendComplete = True
End Sub

Private Sub tmrBuddyStatus_Timer()
   For SeekIndex = 1 To 10
      DoEvents
      If Not Buddy.Buddy(SeekIndex).BuddyNick = "" Then
         SocketSend GETBUDDYSTATUS & DELIMITER & Trim(Buddy.Buddy(SeekIndex).BuddyNick)
      End If
      DoEvents
      DoEvents
   Next
End Sub

Private Sub trvBuddies_DblClick()
   Dim NewIm As New frmIM
   
   If trvBuddies.SelectedItem Is Nothing Then Exit Sub
   If Not trvBuddies.Nodes.Item(trvBuddies.SelectedItem.Index).Bold Then
      NewIm.Caption = Trim(trvBuddies.SelectedItem.Text)
      NewIm.Show
   End If
End Sub

Private Sub trvBuddies_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      Me.PopupMenu frmMenu.mnuUser
   End If
End Sub

Private Sub trvEditBuddies_AfterLabelEdit(Cancel As Integer, NewString As String)
  
  If Trim(NewString) = "" Then
    MsgBox "Item can not be nothing.", vbCritical + vbOKOnly, "Error": Exit Sub
    trvEditBuddies.Nodes.Remove (trvEditBuddies.SelectedItem.Index)
  ElseIf IsValidItem(NewString$) = False Then
    MsgBox "Item can contain only letters, numbers, and spaces.", vbCritical + vbOKOnly, "Error": Exit Sub
    trvEditBuddies.Nodes.Remove (trvEditBuddies.SelectedItem.Index)
  ElseIf ExistsInTree(trvEditBuddies, NewString$) = True Then
    MsgBox Chr(34) & NewString$ & Chr(34) & "Already exists.", vbCritical + vbOKOnly, "Error": Exit Sub
    trvEditBuddies.Nodes.Remove (trvEditBuddies.SelectedItem.Index)
  Else
    If Not trvEditBuddies.SelectedItem.Parent Is Nothing Then
      Buddy.BuddyCount = Buddy.BuddyCount + 1
      
      If Buddy.BuddyCount = 10 Then MsgBox "Too many buddies", vbInformation, "Buddies": Exit Sub
      
      For SeekIndex = 1 To 10
         If Buddy.Buddy(SeekIndex).BuddyNick = "" Then
            Buddy.Buddy(SeekIndex).BuddyNick = NewString
            SaveBuddies NewString
            Exit Sub
         End If
      Next
      
    Else
      If ExistsInTree(trvBuddies, trvEditBuddies.SelectedItem.Text, False, False, NewString$) = False Then
         Set nodGroup = trvBuddies.Nodes.Add(, , , NewString$, 1, 1)
         nodGroup.Bold = True
         nodGroup.Expanded = True
         'savebuddies newstring, true 'means parent [optional]
      End If
   End If
  End If
End Sub

Private Sub trvEditBuddies_DblClick()
  If trvEditBuddies.SelectedItem Is Nothing Then Exit Sub
  trvEditBuddies.StartLabelEdit
End Sub

