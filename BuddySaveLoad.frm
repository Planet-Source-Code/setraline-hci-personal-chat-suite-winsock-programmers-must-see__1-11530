VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Delete Buddy"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add Buddy"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TreeView2 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2143
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2143
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TreeNode As Node

Private Type Buddies2
   BuddyNick As String
   BuddyStat As String
   BuddyInfo As String
End Type

Private Type Buddies
   Buddy(1 To 10) As Buddies2
End Type

Private Buddy As Buddies
Private MaxBuddy As Boolean

Private Sub Command1_Click()
Dim BuddyIndex As Integer

Open "c:\windows\desktop\BuddyList.lst" For Output As #1

Print #1, "Forem"

For BuddyIndex = 1 To 10
   If Not Buddy.Buddy(BuddyIndex).BuddyNick = "" Then
      Print #1, vbTab & Buddy.Buddy(BuddyIndex).BuddyNick
   End If
Next

Close #1
End Sub
' max_user
Private Sub Command2_Click()
LoadTreeViewFromFile "c:\windows\desktop\BuddyList.lst", TreeView2
End Sub

Private Sub Command3_Click()
Dim BuddyIndex As Integer

TreeView1.Nodes.Add TreeNode, tvwChild, , "Ktulu"

For BuddyIndex = 1 To 10
   If Buddy.Buddy(BuddyIndex).BuddyNick = "" Then
      Buddy.Buddy(BuddyIndex).BuddyNick = "Ktulu"
      Exit Sub
   End If
Next

MaxBuddy = True

End Sub

Private Sub Command4_Click()
Dim TreeItem As Integer
Dim BuddyIndex As Integer
Dim BuddyName As String

TreeItem = TreeView1.SelectedItem.Index
BuddyName = TreeView1.Nodes.Item(TreeItem).Text
TreeView1.Nodes.Remove (TreeItem)

For BuddyIndex = 1 To 10
   If Buddy.Buddy(BuddyIndex).BuddyNick = BuddyName Then
      Buddy.Buddy(BuddyIndex).BuddyNick = ""
      Exit For
   End If
Next

End Sub

Private Sub Form_Load()
Dim BuddyIndex As Integer

Set TreeNode = TreeView1.Nodes.Add(, , "Group", "Forem")
TreeNode.Expanded = True
TreeNode.Bold = True
TreeView1.Nodes.Add TreeNode, tvwChild, , "Setraline"

For BuddyIndex = 1 To 10
   If Buddy.Buddy(BuddyIndex).BuddyNick = "" Then
      Buddy.Buddy(BuddyIndex).BuddyNick = "Setraline"
      Exit For
   End If
Next
End Sub

Private Sub LoadTreeViewFromFile(ByVal file_name As String, ByVal trv As TreeView)
    
Dim fnum As Integer
Dim text_line As String
Dim level As Integer
Dim tree_nodes() As Node
Dim num_nodes As Integer

fnum = FreeFile
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
               tree_nodes(level).Bold = True
        Else
        
               Set tree_nodes(level) = trv.Nodes.Add(tree_nodes(level - 1), tvwChild, , text_line)
               'tree_nodes(level).ExpandedImage = "OpenFolder"
               'tree_nodes(level).Image = "ClosedFolder"
               tree_nodes(level).EnsureVisible
        End If
        
    Loop

    Close fnum
    
    Dim BuddyIndex As Integer
    
End Sub

