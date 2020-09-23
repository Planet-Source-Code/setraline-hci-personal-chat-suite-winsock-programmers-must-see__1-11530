VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "IM"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   Picture         =   "frmIM.frx":0000
   ScaleHeight     =   298
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckIM 
      Left            =   4680
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   AutoFormShape Me, vbWhite
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim ReturnVal As Long
   ReleaseCapture
   ReturnVal = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub
