VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "New Folder"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   40
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   40
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   3000
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If File1.ListIndex >= 0 Then
If Right$(Dir1.path, 1) = "\" Then
tmpDir$ = Dir1.path & File1.FileName
Else
tmpDir$ = Dir1.path & "\" & File1.FileName
End If
r = ShellExecute(GetDesktopWindow, "open", tmpDir$, "", Drive1.Drive & "\", 5)
Else
Exit Sub
End If

End Sub

Private Sub Command2_Click()
If File1.ListIndex >= 0 Then
delFile.Show
Exit Sub
End If

File_Delete Dir1.list(Dir1.ListIndex)
Dir1.Refresh
End Sub

Private Sub Command3_Click()
Dim secAtts As SECURITY_ATTRIBUTES
inptReply$ = InputBox("Type the name of the new folder", "New Folder")
If Right$(Dir1.path, 1) = "\" Then
tmpDir$ = Dir1.path & inptReply$
Else
tmpDir$ = Dir1.path & "\" & inptReply$
End If
MsgBox tmpDir$
CreateDirectory tmpDir$, secAtts
Dir1.Refresh

End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
Dir1.path = Drive1.Drive
End Sub

