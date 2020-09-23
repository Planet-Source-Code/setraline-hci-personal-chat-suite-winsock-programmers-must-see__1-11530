VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IM"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   FillColor       =   &H8000000F&
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
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMessage 
      Caption         =   "Message:"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   2640
      Width           =   5895
      Begin RichTextLib.RichTextBox rtbMessage 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   582
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         TextRTF         =   $"frmIMClient.frx":0000
      End
   End
   Begin VB.Frame fraIM 
      Caption         =   "IM"
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5895
      Begin RichTextLib.RichTextBox rtbIM 
         Height          =   1770
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3122
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmIMClient.frx":00EC
      End
   End
   Begin MSComctlLib.ImageList imlFontOptions 
      Left            =   4800
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
            Picture         =   "frmIMClient.frx":01D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   5655
      _ExtentX        =   9975
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
   Begin MSComctlLib.ProgressBar prbProgress 
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   AddSysMessage rtbIM, "*** Forem 2000 IM ***", vbBlue
End Sub

Private Sub rtbMessage_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      KeyAscii = 0
      AddChat rtbIM, ClientNick, rtbMessage.Text, vbBlack
      SendIm Me.Caption, ClientNick, rtbMessage.Text
      rtbMessage.Text = ""
   End If
End Sub
