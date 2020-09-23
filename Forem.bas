Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public ClientUsers() As String
Public SockInUse As Integer

Public Function ParseString(Data As String, Divider As String, SegmentNumber As Integer) As String
   Dim Index As Integer
   Dim Temp As String
   Dim Position As Integer

   Temp = Data

   For Index = 1 To SegmentNumber - 1
      Position = InStr(Temp, Divider)
      If Position Then
         Temp = Mid(Temp, Position + 1)
      Else
         Exit Function
      End If
   Next

   Position = InStr(Temp, Divider)

   If Position Then
      ParseString = Left(Temp, Position - 1)
   Else
      ParseString = Temp
   End If
End Function

Public Sub StopFlicker(ByVal HWnd As Long)
   Dim ReturnVal As Long
   
   ReturnVal = LockWindowUpdate(HWnd)
End Sub

Public Sub Release()
   Dim ReturnVal As Long
    
   ReturnVal = LockWindowUpdate(0)
End Sub

Public Sub AddSysMessage(rtb As RichTextBox, Message As String, Optional ServerColor As Long = vbRed)
   If rtb.Text = "" Then
      Dim NewMessage As String
      NewMessage = Message
   Else
      NewMessage = vbCrLf & Message
   End If
   
      rtb.SelStart = Len(rtb.Text)
      rtb.SelLength = 0
      rtb.SelText = NewMessage
      rtb.SelStart = Len(rtb.Text) - Len(NewMessage)
      rtb.SelLength = Len(NewMessage)
      rtb.SelColor = ServerColor
      rtb.SelBold = False
      rtb.SelFontName = "Courier New"
      rtb.SelFontSize = 10
      rtb.SelStart = Len(rtb.Text)
      rtb.SelLength = 0
End Sub
