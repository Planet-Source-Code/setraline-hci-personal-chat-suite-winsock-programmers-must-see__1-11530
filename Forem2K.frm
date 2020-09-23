VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Forem 2000"
   ClientHeight    =   4485
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6660
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
   Picture         =   "Forem2K.frx":0000
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckConnect 
      Index           =   0
      Left            =   6120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlMenuIcons 
      Left            =   6000
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
            Picture         =   "Forem2K.frx":5F682
            Key             =   ""
            Object.Tag             =   "Send File"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forem2K.frx":5F7DE
            Key             =   ""
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forem2K.frx":61F92
            Key             =   ""
            Object.Tag             =   "Network"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forem2K.frx":64746
            Key             =   ""
            Object.Tag             =   "Kick User"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forem2K.frx":66EFA
            Key             =   ""
            Object.Tag             =   "Get File"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Forem2K.frx":696AE
            Key             =   "Send File"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvUsers 
      Height          =   2235
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "this will show user info"
      Top             =   450
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   3942
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
   Begin VB.TextBox txtIP 
      BackColor       =   &H00808080&
      Height          =   270
      Left            =   600
      MaxLength       =   15
      TabIndex        =   8
      Text            =   "localhost"
      Top             =   2870
      Width           =   2130
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00808080&
      Height          =   270
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "2000"
      Top             =   2870
      Width           =   735
   End
   Begin VB.TextBox txtNick 
      BackColor       =   &H00808080&
      Height          =   270
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   6
      Text            =   "NickName"
      Top             =   2870
      Width           =   1335
   End
   Begin VB.OptionButton optServerClient 
      BackColor       =   &H00808080&
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Client"
      Top             =   3840
      Width           =   255
   End
   Begin VB.OptionButton optServerClient 
      BackColor       =   &H00808080&
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Server"
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   2820
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox rtbMessage 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3420
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   873
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      MaxLength       =   200
      Appearance      =   0
      TextRTF         =   $"Forem2K.frx":6BC42
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2175
      Left            =   495
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3836
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Forem2K.frx":6BD11
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
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Offline"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3240
      TabIndex        =   9
      Top             =   3195
      Width           =   540
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   AutoFormShape Me, vbWhite
   mCoolMenu.Install frmMenu.hwnd, , imlMenuIcons, True, True
   SetTVBackColour vbBlack
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim ReturnVal As Long
   ReleaseCapture
   ReturnVal = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub trvUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then
      Me.PopupMenu frmMenu.mnuList
   End If
End Sub
