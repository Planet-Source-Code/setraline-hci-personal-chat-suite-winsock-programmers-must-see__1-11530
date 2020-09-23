VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forem 2000"
   ClientHeight    =   4785
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8055
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraUsers 
      Caption         =   "People: 0"
      Height          =   3495
      Left            =   6285
      TabIndex        =   8
      Top             =   0
      Width           =   1695
      Begin MSComctlLib.TreeView trvUsers 
         Height          =   3135
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "this will show user info"
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   5530
         _Version        =   393217
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   1
         SingleSel       =   -1  'True
         ImageList       =   "imlMenuIcons"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraMessage 
      Caption         =   "Message:"
      Height          =   735
      Left            =   45
      TabIndex        =   7
      Top             =   3960
      Width           =   6255
      Begin RichTextLib.RichTextBox rtbMessage 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   16777215
         MaxLength       =   150
         TextRTF         =   $"frmChatClient.frx":0000
      End
   End
   Begin VB.Frame fraChat 
      Caption         =   "Chat"
      Height          =   3495
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   6255
      Begin RichTextLib.RichTextBox rtbChat 
         Height          =   3090
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5450
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmChatClient.frx":00D3
      End
   End
   Begin MSComctlLib.ImageList imlFontOptions 
      Left            =   3960
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatClient.frx":01A6
            Key             =   "IMG1"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraChatOptions 
      Caption         =   "Chat Options"
      Height          =   1215
      Left            =   6285
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
      Begin MSComctlLib.ProgressBar prbProgress 
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.ImageList imlMenuIcons 
      Left            =   4560
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatClient.frx":05FA
            Key             =   ""
            Object.Tag             =   "Send File"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatClient.frx":0756
            Key             =   ""
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatClient.frx":2F0A
            Key             =   ""
            Object.Tag             =   "Network"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatClient.frx":56BE
            Key             =   ""
            Object.Tag             =   "Kick User"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatClient.frx":7E72
            Key             =   ""
            Object.Tag             =   "Get File"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChatClient.frx":A626
            Key             =   "Send File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   360
      Left            =   165
      TabIndex        =   1
      Top             =   3600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlFontOptions"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Send File"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Get File"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   rtbMessage.SelBold = False
   rtbMessage.SelUnderline = False
   rtbMessage.SelItalic = False
   rtbMessage.SelColor = vbBlack
   rtbMessage.SelFontName = "Verdana"
   rtbMessage.SelFontSize = 7
   AddSysMessage rtbChat, "*** Joined " & ActiveChannel & " ***", vbRed
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmBuddy.WindowState = vbNormal
   SocketSend CHATLEAVE & DELIMITER & ClientNick
End Sub

Private Sub rtbMessage_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         KeyAscii = 0
         If Not rtbMessage.Text <> "" Then Exit Sub
         Dim ChatCommand As String
      
         ChatCommand = LCase(Trim(ParseString(rtbMessage.Text, "-", 1)))
         Select Case ChatCommand
            Case "/private":
               Dim ToNick As String
               Dim ParsedMessage As String
            
               ToNick = Trim(ParseString(rtbMessage.Text, "-", 2))
               ParsedMessage = Trim(ParseString(rtbMessage.Text, "-", 3))
               AddSysMessage rtbChat, "*** Private Message [ " & ToNick & " ] ***", vbRed
               rtbMessage.Text = ParsedMessage
               AddChat rtbChat, ClientNick, rtbMessage.Text, vbBlack
               SocketSend PRIVATEMESSAGE & DELIMITER & ToNick & DELIMITER & ClientNick & DELIMITER & rtbMessage.Text
               DoEvents
               rtbMessage.Text = ""
               Exit Sub
               AddSysMessage frmChat.rtbChat, "*** Nick Not Found ***", vbRed
            
            Case "/kick":
               Dim KickNick As String
               Dim ReasonKick As String
               
               KickNick = Trim(ParseString(rtbMessage.Text, "-", 2))
               ReasonKick = Trim(ParseString(rtbMessage.Text, "-", 3))
               
               SocketSend KICKUSER & DELIMITER & KickNick & DELIMITER & ReasonKick
               AddSysMessage frmChat.rtbChat, "*** Kicked " & KickNick & " [ " & ReasonKick & " ] ***", vbRed
               
            Case "/afk":
               Dim ReasonAFK As String
               
               ReasonAFK = Trim(ParseString(ChatCommand, DELIMITER, 2))
               SocketSend AFK & DELIMITER & ClientNick & DELIMITER & ReasonAFK
               AddSysMessage frmChat.rtbChat, "*** AFK ***", vbRed
            
            Case Else:
               AddChat rtbChat, ClientNick, rtbMessage.Text, vbBlack
               SocketSend CHATMESSAGE & DELIMITER & ClientNick & DELIMITER & rtbMessage.Text
         End Select
            rtbMessage.Text = ""
   End If
End Sub

Private Sub rtbMessage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   rtbMessage.SelBold = False
   rtbMessage.SelUnderline = False
   rtbMessage.SelItalic = False
   rtbMessage.SelColor = vbBlack
   rtbMessage.SelFontName = "Verdana"
   rtbMessage.SelFontSize = 7
End Sub

Private Sub rtbMessage_SelChange()
   rtbMessage.SelBold = False
   rtbMessage.SelUnderline = False
   rtbMessage.SelItalic = False
   rtbMessage.SelColor = vbBlack
   rtbMessage.SelFontName = "Verdana"
   rtbMessage.SelFontSize = 7
End Sub

Private Sub trvUsers_DblClick()
   Dim NewIm As New frmIM
   
   NewIm.Caption = Trim(trvUsers.SelectedItem.Text)
   NewIm.Show
End Sub

Private Sub trvUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      Me.PopupMenu frmMenu.mnuUser
   End If
End Sub

