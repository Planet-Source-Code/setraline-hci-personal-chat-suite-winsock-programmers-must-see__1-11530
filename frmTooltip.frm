VERSION 5.00
Begin VB.Form frmTooltip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleWidth      =   1440
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrParent 
      Interval        =   5
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Left            =   120
      Top             =   0
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "frmTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API declares
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

' constants
Private Const OffSetX = -2
Private Const OffSetY = 18
Private Const HWND_TOP& = 0
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOACTIVATE& = &H10
Private Const SWP_NOSIZE& = &H1
Private Const SWP_SHOWWINDOW& = &H40

' types
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' member variables
Private m_Margin As Integer
Private m_hWnd As Long
Private m_End As Boolean

Private Sub Form_Load()

    Me.Margin = 3
    Me.ShowTime = 3000

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Me.Hide
    
End Sub

Private Sub Form_Resize()
    
    Call Cls
    Me.ScaleMode = 3
    Line (0, 0)-(Me.ScaleWidth, 0), vb3DLight
    Line (0, 0)-(0, Me.ScaleHeight), vb3DLight
    Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight), vb3DDKShadow
    Line (0, Me.ScaleHeight - 1)-(Me.ScaleWidth, Me.ScaleHeight - 1), vb3DDKShadow
    Me.ScaleMode = 1
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call Me.Hide
    If m_End Then
        Cancel = 0
    Else
        Cancel = 1
    End If

End Sub

Private Sub lblTip_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call Me.Hide
 
End Sub

Public Property Let Text(sText As String)

    lblTip.Caption = sText
        
End Property

Public Property Get Text() As String

    Text = lblTip.Caption

End Property

Public Property Let Margin(sNum As Integer)
    
    m_Margin = sNum
    Call lblTip.Move(m_Margin * Screen.TwipsPerPixelX, m_Margin * Screen.TwipsPerPixelY)
    
End Property

Public Property Get Margin() As Integer

    Margin = m_Margin

End Property

Public Property Let ToolFont(sFont As String)

    lblTip.Font = sFont

End Property

Public Property Get ToolFont() As String

    Font = lblTip.Font

End Property

Public Property Let ToolForeColor(lColor As Long)

    lblTip.ForeColor = lColor

End Property

Public Property Get ToolForeColor() As Long

    ToolForeColor = lblTip.ForeColor

End Property

Public Property Let ToolBackColor(lColor As Long)

    Me.BackColor = lColor

End Property

Public Property Get ToolBackColor() As Long

    ToolBackColor = Me.BackColor

End Property

Public Property Let ShowTime(nMilliseconds As Long)

    tmrShow.Interval = nMilliseconds

End Property

Public Property Get ShowTime() As Long

    ShowTime = tmrShow.Interval

End Property

Public Sub ShowTip()
    
    If Me.Visible Then Exit Sub
    If lblTip.Caption = "" Then Exit Sub
    Call UpdatePos
    If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
    If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
    Call SetWindowPos(Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW)
    tmrShow.Enabled = True
    tmrParent.Enabled = True
    
End Sub

Private Sub tmrParent_Timer()

    Dim ptCursor As POINTAPI
    Dim rWindow As RECT
    
    Call GetCursorPos(ptCursor)
    Call GetWindowRect(m_hWnd, rWindow)
    
    If ptCursor.X < rWindow.Left Or ptCursor.X > rWindow.Right Or ptCursor.Y < rWindow.Top Or ptCursor.Y > rWindow.Bottom Then Call Unload(Me)

End Sub

Private Sub tmrShow_Timer()

    Call Me.Hide

End Sub

Public Sub CreateTip(obj As Object)

    On Error GoTo Cancel
    m_hWnd = obj.hwnd

Cancel:
End Sub

Public Sub HideTip()

    Call Me.Hide

End Sub

Public Sub UpdatePos()

    Dim ptCursor As POINTAPI

    Call GetCursorPos(ptCursor)
    Call Me.Move((ptCursor.X + OffSetX) * Screen.TwipsPerPixelX, (ptCursor.Y + OffSetY) * Screen.TwipsPerPixelY, lblTip.Width + 2 * (m_Margin * Screen.TwipsPerPixelX), lblTip.Height + 2 * (m_Margin * Screen.TwipsPerPixelY))
    
End Sub

Public Sub DestroyTip()

    m_End = True
    Call Form_Unload(0)

End Sub

Public Property Let ToolFontBold(bBold As Boolean)

    lblTip.FontBold = bBold

End Property

Public Property Get ToolFontBold() As Boolean

    ToolFontBold = lblTip.FontBold

End Property

Public Property Let ToolFontItalic(bItalic As Boolean)

    lblTip.FontItalic = bItalic

End Property

Public Property Get ToolFontItalic() As Boolean

    ToolFontItalic = lblTip.FontItalic

End Property

Public Property Let ToolFontSize(nFont As Single)

    lblTip.FontSize = nFont

End Property

Public Property Get ToolFontSize() As Single

    ToolFontSize = lblTip.FontSize

End Property
