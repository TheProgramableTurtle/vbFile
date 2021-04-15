VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VBFile"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120
   Icon            =   "vbFile.frx":0000
   LinkTopic       =   "frmMain"
   ScaleHeight     =   6105
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInfo 
      Caption         =   "About"
      Height          =   495
      Left            =   6960
      TabIndex        =   14
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8040
      TabIndex        =   13
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New File"
      Height          =   495
      Left            =   3360
      TabIndex        =   12
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   405
      Left            =   2640
      TabIndex        =   3
      Top             =   5040
      Width           =   5535
   End
   Begin VB.DriveListBox drvMain 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.FileListBox fleMain 
      Height          =   4185
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.DirListBox dirMain 
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label lblFlePath 
      Caption         =   "File path"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblDrv 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblFle 
      Caption         =   "Files"
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblDir 
      Caption         =   "Directories"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()
    If MsgBox("Permanently delete " & fleMain.List(fleMain.ListIndex) & "?", vbOKCancel + vbQuestion, "Delete file?") = vbOK Then
        Kill fleMain.Path & "\" & fleMain.List(fleMain.ListIndex)
    End If
    fleMain.Refresh
    dirMain.Refresh
End Sub

Private Sub cmdEdit_Click()
    Shell "notepad.exe " & (fleMain.Path & "\" & fleMain.List(fleMain.ListIndex)), vbNormalFocus
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGo_Click()
    On Error Resume Next
    drvMain.Drive = Mid(dirMain.Path, 1, 2)
    dirMain.Path = txtPath.Text
    txtPath.Text = fleMain.Path
End Sub

Private Sub cmdInfo_Click()
    If MsgBox("VBFile" & vbCrLf & "v2.0" & vbCrLf & "©Pr0x1mas 2021", vbInformation, "About") = vbOK Then
        ' do nothing
    End If
     
End Sub

Private Sub cmdNew_Click()
    Open fleMain.Path & "\" & InputBox("Enter filename", "New File") For Output As #1
    fleMain.Refresh
    dirMain.Refresh
    Close #1
End Sub

Private Sub cmdOpen_Click()
    Shell "explorer.exe " & (fleMain.Path & "\" & fleMain.List(fleMain.ListIndex)), vbNormalFocus
End Sub

Private Sub dirMain_Change()
    fleMain.Path = dirMain.Path
    txtPath.Text = fleMain.Path
End Sub

Private Sub drvMain_Change()
    On Error Resume Next
    dirMain.Path = drvMain.Drive
    drvMain.Drive = Mid(dirMain.Path, 1, 2)
End Sub

Private Sub Form_Load()
    txtPath.Text = fleMain.Path

End Sub

