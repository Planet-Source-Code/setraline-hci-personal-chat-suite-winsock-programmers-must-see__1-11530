VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Message Box"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProceed 
      BackColor       =   &H8000000C&
      Caption         =   "&Proceed"
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
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image imgIcon 
      Height          =   495
      Index           =   3
      Left            =   2280
      Picture         =   "frmMessageBox.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   495
      Index           =   2
      Left            =   1560
      Picture         =   "frmMessageBox.frx":0742
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   495
      Index           =   1
      Left            =   840
      Picture         =   "frmMessageBox.frx":0E84
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image imgIcon 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frmMessageBox.frx":12C6
      Stretch         =   -1  'True
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProceed_Click()
   Me.Hide
End Sub

