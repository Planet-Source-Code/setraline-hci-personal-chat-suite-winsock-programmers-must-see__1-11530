VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Forem 2000"
   ClientHeight    =   4290
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChat.frx":0000
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   498
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin MSWinsockLib.Winsock sckChat 
      Left            =   5040
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlMenuIcons 
      Left            =   4920
      Top             =   360
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
            Picture         =   "frmChat.frx":6719E
            Key             =   ""
            Object.Tag             =   "Send File"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":672FA
            Key             =   ""
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":69AAE
            Key             =   ""
            Object.Tag             =   "Network"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":6C262
            Key             =   ""
            Object.Tag             =   "Kick User"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":6EA16
            Key             =   ""
            Object.Tag             =   "Get File"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":711CA
            Key             =   "Send File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvUsers 
      Height          =   2295
      Left            =   5760
      TabIndex        =   2
      ToolTipText     =   "this will show user info"
      Top             =   405
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   4048
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   1
      SingleSel       =   -1  'True
      ImageList       =   "imlMenuIcons"
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
   Begin RichTextLib.RichTextBox rtbMessage 
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   3420
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      MaxLength       =   200
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":7375E
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2175
      Left            =   495
      TabIndex        =   1
      Top             =   480
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":73848
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   Me.Hide
End Sub

Private Sub Form_Load()
   'AutoFormShape Me, vbWhite
   SetTVBackColour vbBlack
   rtbMessage.SelBold = False
   rtbMessage.SelUnderline = False
   rtbMessage.SelItalic = False
   rtbMessage.SelColor = vbYellow
   rtbMessage.SelFontName = "Verdana"
   rtbMessage.SelFontSize = 7
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim ReturnVal As Long
   
   ReleaseCapture
   ReturnVal = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub rtbMessage_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         KeyAscii = 0
         AddChat rtbChat, "Lee", rtbMessage.TextRTF
         rtbMessage.Text = ""
   End If
End Sub

Private Sub rtbMessage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   rtbMessage.SelBold = False
   rtbMessage.SelUnderline = False
   rtbMessage.SelItalic = False
   rtbMessage.SelColor = vbYellow
   rtbMessage.SelFontName = "Verdana"
   rtbMessage.SelFontSize = 7
End Sub

Private Sub rtbMessage_SelChange()
   rtbMessage.SelBold = False
   rtbMessage.SelUnderline = False
   rtbMessage.SelItalic = False
   rtbMessage.SelColor = vbYellow
   rtbMessage.SelFontName = "Verdana"
   rtbMessage.SelFontSize = 7
End Sub

Private Sub trvUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      Me.PopupMenu frmMenu.mnuList
   End If
End Sub
