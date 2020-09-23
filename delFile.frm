VERSION 5.00
Begin VB.Form delFile 
   Caption         =   "Delete"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form2"
   ScaleHeight     =   1470
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Delete File or Folder?"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   3240
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Delete File:"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   4935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Delete Folder:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "delFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim FileOperation As SHFILEOPSTRUCT
Dim lReturn As Long
Dim sSendMeToTheBin As String
If Option1.Value = True Then
File_Delete Form1.Dir1.List(Form1.Dir1.ListIndex)
Unload Me
Form1.Dir1.Refresh
Else
    If Right$(Form1.Dir1.path, 1) = "\" Then
    File_Delete Form1.Dir1.path & Form1.File1.FileName
    Unload Me
    Form1.File1.Refresh
    Else
    File_Delete Form1.Dir1.path & "\" & Form1.File1.FileName
    Unload Me
    Form1.File1.Refresh
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Option1.Caption = "Delete Folder: " & Form1.Dir1.List(Form1.Dir1.ListIndex)
If Right$(Form1.Dir1.path, 1) = "\" Then
Option2.Caption = "Delete File: " & Form1.Dir1.path & Form1.File1.FileName
Else
Option2.Caption = "Delete File: " & Form1.Dir1.path & "\" & Form1.File1.FileName
End If
End Sub
