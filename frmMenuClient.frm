VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   135
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1680
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Small Fonts"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   135
   ScaleWidth      =   1680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlMenus 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuClient.frx":0000
            Key             =   ""
            Object.Tag             =   "Kick User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuClient.frx":015C
            Key             =   ""
            Object.Tag             =   "Send File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuClient.frx":02B8
            Key             =   ""
            Object.Tag             =   "Get File"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuClient.frx":2A6C
            Key             =   ""
            Object.Tag             =   "Join Chat"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuClient.frx":5220
            Key             =   ""
            Object.Tag             =   "Instant Message"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuClient.frx":79D4
            Key             =   ""
            Object.Tag             =   "Settings"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenuClient.frx":A188
            Key             =   ""
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send File"
      End
      Begin VB.Menu mnuGetFile 
         Caption         =   "Get File"
      End
      Begin VB.Menu mnuJoinChat 
         Caption         =   "Join Chat"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
   mCoolMenu.Uninstall Me.hWnd
End Sub

Private Sub mnuExit_Click()
   mCoolMenu.Uninstall Me.hWnd
   End
End Sub

Private Sub mnuJoinChat_Click()
   Dim ChatName As String
   
   ChatName = InputBox("Enter chat room to join:", "Join Chat", "Forem 2000")
   
   If Not ChatName <> "" Then Exit Sub
   
   SocketSend CHATJOIN & DELIMITER & ChatName & DELIMITER & ClientNick
   ActiveChannel = ChatName
   frmChat.trvUsers.Nodes.Add , , , ClientNick, 1
   frmChat.Show
   frmBuddy.WindowState = vbMinimized
End Sub

Private Sub mnuSendFile_Click()
   'SendFile frmbuddy.scksendfile, thefile
End Sub
