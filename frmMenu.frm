VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   300
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   1215
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlMenus 
      Left            =   360
      Top             =   0
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
            Picture         =   "frmMenu.frx":0000
            Key             =   ""
            Object.Tag             =   "Kick User"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":27B4
            Key             =   ""
            Object.Tag             =   "Send File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3388
            Key             =   ""
            Object.Tag             =   "Get File"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":34E4
            Key             =   ""
            Object.Tag             =   "Exit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":844C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Begin VB.Menu mnuKickUser 
         Caption         =   "Kick User"
      End
      Begin VB.Menu mnuSendFile 
         Caption         =   "Send File"
      End
      Begin VB.Menu mnuGetFile 
         Caption         =   "Get File"
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

Private Sub mnuExit_Click()
   mCoolMenu.Uninstall Me.hWnd
   End
End Sub

