VERSION 5.00
Begin VB.Form frmSignOn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sign On"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
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
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBorder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2775
      Begin VB.Image imgSplash 
         Height          =   975
         Left            =   120
         Picture         =   "frmSignOn.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSignOn 
      BackColor       =   &H8000000C&
      Caption         =   "Sign On"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Setraline"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox cboScreenName 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Setraline"
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblPassword 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblScreenName 
      Caption         =   "Screen Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frmSignOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSignOn_Click()
   Server = Trim(txtHost)
   ServerPort = 3370
   ClientNick = Trim(cboScreenName.Text)
   
   If SocketConnect(Server, ServerPort) Then
      Play "connect.wav"
      Load frmBuddy
      Me.Hide
      frmBuddy.Show
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   mCoolMenu.Uninstall frmMenu.hWnd
   End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   mCoolMenu.Uninstall frmMenu.hWnd
   End
End Sub

