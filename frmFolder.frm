VERSION 5.00
Begin VB.Form frmFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Folder"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   405
      Left            =   4470
      TabIndex        =   10
      Tag             =   "Confirm"
      Top             =   30
      Width           =   1605
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   4470
      TabIndex        =   9
      Tag             =   "Cancel"
      Top             =   480
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Caption         =   "Special Folders:"
      Height          =   2655
      Left            =   4470
      TabIndex        =   4
      Top             =   1620
      Width           =   1575
      Begin VB.CommandButton cmdDesktop 
         Caption         =   "Desktop"
         Height          =   435
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Tag             =   "Desktop"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDesktop 
         Caption         =   "Documents"
         Height          =   435
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Tag             =   "Documents"
         Top             =   900
         Width           =   975
      End
      Begin VB.CommandButton cmdDesktop 
         Caption         =   "Download"
         Height          =   435
         Index           =   2
         Left            =   270
         TabIndex        =   6
         Tag             =   "Download"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdDesktop 
         Caption         =   "User"
         Height          =   435
         Index           =   3
         Left            =   270
         TabIndex        =   5
         Tag             =   "User"
         Top             =   2010
         Width           =   975
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
   Begin VB.Frame fraFileNames 
      Caption         =   "Folder:"
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   4350
      Width           =   6045
      Begin VB.TextBox txtFilename 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   5865
      End
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sEmuFolder As String
Public isTarget     As Boolean

Private Sub cmdCancel_Click()
    sEmuFolder = ""
    Unload Me
End Sub

'Set Special path on Dir Navigator
Private Sub cmdDesktop_Click(Index As Integer)

On Error GoTo ERR_SUB
    
    Drive1.Drive = Environ("SystemDrive")
    
    Select Case Index
        Case 0
            Dir1.Path = Environ("USERPROFILE") & "\Desktop"
        Case 1
            Dir1.Path = Environ("USERPROFILE") & "\Documents"
        Case 2
            Dir1.Path = Environ("USERPROFILE") & "\Downloads"
        Case 3
            Dir1.Path = Environ("USERPROFILE") & "\"
    End Select
    
    txtFilename.Text = Dir1.Path

    
EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "SD FormLoad Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB
    

End Sub

Private Sub cmdOk_Click()
    sEmuFolder = txtFilename.Text
    Unload Me
End Sub

Private Sub Dir1_Change()
    txtFilename.Text = Dir1.Path
End Sub

Private Sub Dir1_Click()
    txtFilename.Text = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()

On Error GoTo ERR_SUB

    If Trim(sEmuFolder) = "" Then

        
        If isTarget Then
            Drive1.Drive = App.Path
            Dir1.Path = App.Path
        Else
            Drive1.Drive = Environ("USERPROFILE")
            Dir1.Path = Environ("USERPROFILE") & "\Desktop"
        End If
    Else
        Drive1.Drive = sEmuFolder
        Dir1.Path = sEmuFolder
    End If
    
    
    sEmuFolder = Dir1.Path
    txtFilename.Text = Dir1.Path
    
EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "SD FormLoad - You can go on but here there is an Unexpected Error n." & Err.Number & vbCrLf & Err.Description & vbCrLf & "Please report the error if you can"
    Resume EXIT_SUB
    

End Sub
