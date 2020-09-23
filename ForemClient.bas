Attribute VB_Name = "Module2"
Option Explicit

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

Public ChannelName As String
Public ActiveChannel As String
Public SeekIndex As Integer
Public Server As String
Public ServerPort As Long
Public ClientNick As String
Public nodGroup As Node
Public SendComplete As Boolean

Public Type Buddies2
   BuddyNick As String
   BuddyStat As String
   BuddyInfo As String
   BuddyConnectionIndex As Integer
End Type

Public Type Buddies
   Buddy(1 To 10) As Buddies2
   BuddyCount As Integer
End Type

Public Buddy As Buddies

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

Public Sub StopFlicker(ByVal hWnd As Long)
   Dim ReturnVal As Long
   
   ReturnVal = LockWindowUpdate(hWnd)
End Sub

Public Sub Release()
   Dim ReturnVal As Long
    
   ReturnVal = LockWindowUpdate(0)
End Sub

Public Sub AddChat(rtb As RichTextBox, Nick As String, Message As String, ColorCode As Long)
   rtb.SelStart = Len(rtb.Text)
   rtb.SelBold = True
   rtb.SelColor = vbBlue
   rtb.SelText = IIf(rtb.Text = "", Nick & ":", vbCrLf & Nick & ":")
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

Public Sub SendFile(sck As Winsock, SendFileName As String)
   Dim CheckSum As Long
   Dim FileData() As Byte
   Dim FileHandle As Long


          frmBuddy.cdgFileDialog.Flags = cdlOFNHideReadOnly
          On Error GoTo CancelClicked
          frmBuddy.cdgFileDialog.ShowOpen

          If frmBuddy.cdgFileDialog.FileName <> "" Then
          
               SendFileName = frmBuddy.cdgFileDialog.FileName
               frmBuddy.cdgFileDialog.FileName = ""
               
          End If

          DoEvents

          FileHandle = FreeFile
          
          Open SendFileName For Binary As FileHandle
          ReDim FileData(0 To LOF(FileHandle) - 1)
          Get FileHandle, , FileData()
          Close FileHandle
          
         ' SendFileName frmBuddy.cdgFileDialog.FileTitle

          If frmBuddy.sckSendFile.State <> sckClosed Then frmBuddy.sckSendFile.Close

          frmBuddy.sckSendFile.Connect "127.0.0.1", ServerPort + 1
          
          'Status "connecting to " & "127.0.0.1" & "..."

          Do Until frmBuddy.sckSendFile.State = sckConnected Or frmBuddy.sckSendFile.State = sckError
               DoEvents
          Loop

          If frmBuddy.sckSendFile.State = sckError Then Exit Sub

          'lblCaption = "live wire v1.01 [ " & frmbuddy.frmbuddy.scksendfile.RemoteHostIP & " ] - remote"

          'Status "connected to " & "127.0.0.1"
         
          CheckSum = GetArrayCheckSum(FileData)

          'frmBuddy.frmBuddy.sckSendFile.SendData FileName & DELIMITER & CStr(UBound(FileData) + 1) & DELIMITER & CStr(CheckSum) & DELIMITER
CancelClicked:

End Sub

Public Function GetArrayCheckSum(Data() As Byte) As Long

Dim BytePointer As Long
Dim RunningTotal As Long

For BytePointer = 0 To UBound(Data)

    RunningTotal = RunningTotal + Data(BytePointer)
    
Next BytePointer

GetArrayCheckSum = RunningTotal

End Function

Public Function ExistsInTree(tvw As TreeView, ByVal strItem As String, Optional blnStartEdit As Boolean = False, Optional blnDelete As Boolean = False, Optional strReplaceWith As String = "") As Boolean
  'this procedure is used to handle the buddylist treeviews.
  Dim lngDo As Long, blnExists As Boolean
  blnExists = False
  strItem$ = LCase(Replace(strItem$, " ", ""))
  For lngDo& = 1 To tvw.Nodes.Count
    If strItem$ = LCase(Replace(tvw.Nodes.Item(lngDo&).Text, " ", "")) Then
      blnExists = True
      If blnStartEdit = True Then
        tvw.SetFocus
        tvw.Nodes.Item(lngDo&).Selected = True
        tvw.StartLabelEdit
      End If
      If blnDelete = True Then
        tvw.Nodes.Remove lngDo&
      End If
      If strReplaceWith$ <> "" Then
        tvw.Nodes.Item(lngDo&).Text = strReplaceWith$
      End If
      Exit For
    End If
  Next
  ExistsInTree = blnExists
End Function

Public Function IsValidItem(strItem As String) As Boolean
  Dim lngDo As Long, blnIsValid As Boolean, strChar As String
  blnIsValid = True
  For lngDo& = 1 To Len(strItem$)
    strChar$ = Mid(strItem$, lngDo&, 1)
    If Asc(strChar$) < 65 Or Asc(strChar$) > 90 Then
      If Asc(strChar$) < 97 Or Asc(strChar$) > 122 Then
        If IsNumeric(strChar$) = False Then
          If strChar$ <> " " Then
            blnIsValid = False
            Exit For
          End If
        End If
      End If
    End If
  Next
  IsValidItem = blnIsValid
End Function

Public Sub SendIm(ToNick As String, FromNick As String, TheMessage As String)
   SocketSend USERMESSAGE & DELIMITER & ToNick & DELIMITER & FromNick & DELIMITER & TheMessage
End Sub

Public Sub Play(SoundFile As String)
   On Error Resume Next
   sndPlaySound App.Path & "\" & SoundFile, 1
End Sub

Public Sub SocketSend(SocketData As String)
   If frmBuddy.sckClient.State = sckConnected Then
      frmBuddy.sckClient.SendData SocketData
   End If
End Sub

Public Function SocketConnect(SocketIP As String, SocketPort As Long) As Boolean
   frmBuddy.sckClient.Close
   frmBuddy.sckClient.Connect SocketIP, SocketPort
   
   Do Until frmBuddy.sckClient.State = sckConnected Or frmBuddy.sckClient.State = sckError
      DoEvents
   Loop
   
   If frmBuddy.sckClient.State = sckConnected Then
      SocketConnect = True
   Else
      SocketConnect = False
   End If

End Function

Public Sub SaveBuddies(Optional NewBuddy As String)
   Dim BuddyIndex As Integer

   Open App.Path & "\BuddyList.lst" For Output As #1

   Print #1, "Forem"

   If Not NewBuddy = "" Then
      For BuddyIndex = 1 To 10
         If Buddy.Buddy(BuddyIndex).BuddyNick = "" Then
            Buddy.Buddy(BuddyIndex).BuddyNick = NewBuddy
            Exit For
         End If
      Next
   End If
   
   For BuddyIndex = 1 To 10
      If Not Buddy.Buddy(BuddyIndex).BuddyNick = "" Then
         Print #1, vbTab & Buddy.Buddy(BuddyIndex).BuddyNick
      End If
   Next

   Close #1
End Sub

Public Sub LoadBuddies()
   If FileExists(App.Path & "\BuddyList.lst") Then
      LoadTreeViewFromBuddyFile App.Path & "\BuddyList.lst", frmBuddy.trvEditBuddies
      frmBuddy.cmdAddGroup.Enabled = False
   End If
End Sub

Private Sub LoadTreeViewFromFile(ByVal file_name As String, ByVal trv As TreeView)
    
Dim fnum As Integer
Dim text_line As String
Dim level As Integer
Dim tree_nodes() As Node
Dim num_nodes As Integer

fnum = FreeFile
On Error GoTo FileError
Open file_name For Input As fnum

trv.Nodes.Clear
Do While Not EOF(fnum)
     DoEvents
     ' Get a line.
     Line Input #fnum, text_line

     ' Find the level of indentation.
     level = 1
     
     Do While Left$(text_line, 1) = vbTab
          level = level + 1
          text_line = Mid$(text_line, 2)
          DoEvents
     Loop

     ' Make room for the new node.
        If level > num_nodes Then
          num_nodes = level
          ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then
        
               Set tree_nodes(level) = trv.Nodes.Add(, , , text_line)
               'tree_nodes(level).ExpandedImage = "OpenFolder"
               'tree_nodes(level).Image = "ClosedFolder"
        Else
        
               Set tree_nodes(level) = trv.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line)
               'tree_nodes(level).ExpandedImage = "OpenFolder"
               'tree_nodes(level).Image = "ClosedFolder"
               tree_nodes(level).EnsureVisible
               
        End If
        
    Loop
FileError:
    Close fnum
    
End Sub

Private Sub LoadTreeViewFromBuddyFile(ByVal file_name As String, ByVal trv As TreeView)
    
Dim fnum As Integer
Dim text_line As String
Dim level As Integer
Dim tree_nodes() As Node
Dim num_nodes As Integer
Dim BuddyIndex As Integer

fnum = FreeFile
On Error GoTo FileError
Open file_name For Input As fnum

trv.Nodes.Clear
Do While Not EOF(fnum)
     DoEvents
     ' Get a line.
     Line Input #fnum, text_line

     ' Find the level of indentation.
     level = 1
     
     Do While Left$(text_line, 1) = vbTab
          level = level + 1
          text_line = Mid$(text_line, 2)
          DoEvents
     Loop

     ' Make room for the new node.
        If level > num_nodes Then
          num_nodes = level
          ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        If level = 1 Then
        
               Set tree_nodes(level) = trv.Nodes.Add(, , , text_line, 1)
               tree_nodes(level).Bold = True
               tree_nodes(level).Expanded = True
               Set nodGroup = frmBuddy.trvBuddies.Nodes.Add(, , , text_line, 1)
               nodGroup.Bold = True
               nodGroup.Expanded = True
        Else
        
               Set tree_nodes(level) = trv.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line, 3)
               tree_nodes(level).EnsureVisible
               
               For BuddyIndex = 1 To 10
                  If Buddy.Buddy(BuddyIndex).BuddyNick = "" Then
                     Buddy.Buddy(BuddyIndex).BuddyNick = text_line
                  Exit For
                  End If
               Next
      End If
        
    Loop
FileError:
    Close fnum
    
End Sub

Public Function FileExists(strFileName As String) As Boolean
  Dim intLen As Integer
  If strFileName$ <> "" Then
    intLen% = Len(Dir$(strFileName$))
    FileExists = (Not Err And intLen% > 0)
  Else
    FileExists = False
  End If
End Function
