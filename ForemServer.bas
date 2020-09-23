Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SERVERMESSAGE = "1"
Public Const CHATJOIN = "2"
Public Const CHATLEAVE = "3"
Public Const USERMESSAGE = "4"
Public Const NEWUSER = "5"
Public Const KICKUSER = "6"
Public Const NICKINUSE = "7"
Public Const DELIMITER = "*"
Public Const IM = "8"
Public Const USERINFORMATION = "9"
Public Const CHATMESSAGE = "a"
Public Const GETBUDDYSTATUS = "b"
Public Const BUDDYSTATUS = "c"
Public Const ONLINE = "d"
Public Const OFFLINE = "e"
Public Const CHATUSERS = "f"
Public Const PRIVATEMESSAGE = "g"
Public Const AFK = "h"

Public SockInUse As Integer
Public Channel As String
Public SeekIndex As Integer
Public ActiveChannel As String
Public ChannelCount As Integer
Public ServerNick As String
Public MaxUsers As Integer

Public ToolTip As frmTooltip

Public Type UserInfo2
   UserName As String * 10
   UserIndex As Integer
   Connected As Boolean
   info As String * 50
   InChannel As Boolean
   Channel As String * 10
End Type

Public Type UserInfo
   User() As UserInfo2
   UserCount As Integer
End Type

Public User As UserInfo

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

Public Sub StopFlicker(ByVal hwnd As Long)
   Dim ReturnVal As Long
   
   ReturnVal = LockWindowUpdate(hwnd)
End Sub

Public Sub Release()
   Dim ReturnVal As Long
    
   ReturnVal = LockWindowUpdate(0)
End Sub

Public Sub AddChat(rtb As RichTextBox, NIck As String, Message As String, ColorCode As Long)
   rtb.SelStart = Len(rtb.Text)
   rtb.SelBold = True
   rtb.SelColor = vbBlue
   rtb.SelText = IIf(rtb.Text = "", NIck & ":", vbCrLf & NIck & ":")
   rtb.SelStart = Len(rtb.Text)
   rtb.SelColor = vbWhite
   rtb.SelBold = False
   rtb.SelText = "__"
   rtb.SelStart = Len(rtb.Text)
   rtb.SelBold = False
   rtb.SelColor = ColorCode
   rtb.SelHangingIndent = 300
   rtb.SelText = Message
End Sub

Public Sub AddSysMessage(rtb As RichTextBox, Message As String, ColorCode As Long)
   rtb.SelStart = Len(rtb.Text)
   rtb.SelBold = True
   rtb.SelColor = ColorCode
   rtb.SelHangingIndent = 300
   rtb.SelText = IIf(rtb.Text = "", Message, vbCrLf & Message)
End Sub

Public Sub SendIm(ToNick As String, FromNick As String, TheMessage As String, ImIndex As Integer)
   SocketSend ImIndex, USERMESSAGE & DELIMITER & ToNick & DELIMITER & FromNick & DELIMITER & TheMessage
End Sub

Public Sub Play(SoundFile As String)
   On Error Resume Next
   sndPlaySound App.Path & "\" & SoundFile, 1
End Sub

Public Sub SocketSend(SocketID As Integer, SocketData As String)
   If frmMain.sckServer(SocketID).State = sckConnected Then
      frmMain.sckServer(SocketID).SendData SocketData
   End If
End Sub


