VERSION 5.00
Begin VB.Form frmSD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Sd Card"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "FrmFolder"
   Begin VB.ComboBox cboDrives 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   420
      Width           =   3765
   End
   Begin VB.Frame fraSectorNo 
      Appearance      =   0  'Flat
      Caption         =   "Sector:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   4110
      TabIndex        =   5
      Tag             =   "Sector"
      Top             =   1020
      Width           =   1605
      Begin VB.TextBox txtSector 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Text            =   "0x"
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   4110
      TabIndex        =   4
      Tag             =   "Cancel"
      Top             =   510
      Width           =   1605
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   405
      Left            =   4110
      TabIndex        =   3
      Tag             =   "Confirm"
      Top             =   60
      Width           =   1605
   End
   Begin VB.CommandButton cmdRefresh 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   345
      Left            =   30
      TabIndex        =   2
      Tag             =   "Refresh"
      Top             =   30
      Width           =   1875
   End
   Begin VB.CheckBox chkAllDrive 
      Caption         =   "Show All Drives"
      Height          =   285
      Left            =   2100
      TabIndex        =   1
      Tag             =   "ShowAllDrive"
      Top             =   60
      Width           =   1665
   End
   Begin VB.ListBox lstDrives 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "frmSD.frx":0000
      Left            =   30
      List            =   "frmSD.frx":0002
      TabIndex        =   0
      Top             =   780
      Width           =   3765
   End
End
Attribute VB_Name = "frmSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oDrive As clsDrive
Public isTarget As Boolean
Public enTipo   As enEmuType

Private m_bLoadingForm As Boolean


Private Sub cboDrives_Change()
   Dim oDrv As clsDrive
   
   'Select
    If Not g_cDrives Is Nothing Then
        
        For Each oDrv In g_cDrives
        
            If oDrv.drvIndex = cboDrives.ListIndex Then
                Set oDrive = oDrv
                Exit For
            End If
            
        Next
        txtSector.Text = "0x" & CStr(Hex(oDrive.drvStartSector * 10000@))
        
    End If
End Sub

Private Sub cboDrives_Click()
   Dim oDrv As clsDrive
   
   'Select
    If Not g_cDrives Is Nothing Then
        
        For Each oDrv In g_cDrives
        
            If oDrv.drvIndex = cboDrives.ListIndex Then
                Set oDrive = oDrv
                Exit For
            End If
            
        Next
        
        If Not isTarget Then
        
            txtSector.Text = "0x" & CStr(Hex(oDrive.drvStartSector * 10000@))
        End If
    
        If enTipo = enAtmHidden Then
            lstDrives.Visible = True
            FillList
        Else
            lstDrives.Visible = False
            txtSector.Text = "0x2"
        End If

    End If
    
End Sub

Private Sub chkAllDrive_Click()

    If Not m_bLoadingForm Then
    
        On Error GoTo ERR_SUB
    
        If chkAllDrive.Value = vbChecked Then
            If MsgBox(LangLabel("MSGShowAllDrives", "Are you sure you want to show all drive?" & vbCrLf & "Be very, very careful on what you are doing"), vbYesNo + vbQuestion, LangLabel("Warning", "Warning")) = vbYes Then
                g_bFindAllDrives = True
                cmdRefresh_Click
            End If
        Else
            g_bFindAllDrives = False
            cmdRefresh_Click
        End If
    End If

EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "chkAllDrive Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB

End Sub

Private Sub cmdCancel_Click()
    Set oDrive = Nothing
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    Dim oDrv As clsDrive
    Dim sTmp As String

On Error GoTo EXIT_SUB

    'Select
    If Not g_cDrives Is Nothing Then
        
        For Each oDrv In g_cDrives
        
            If oDrv.drvIndex = cboDrives.ListIndex Then
                Set oDrive = oDrv
                Exit For
            End If
            
        Next
        
'        If oDrive.drvStartSector > 0 Then txtSector.Text = "0x" & CStr(Hex(oDrive.drvStartSector * 10000@))
        
    End If
    
    'Retrieve the text field value
    sTmp = Replace(txtSector.Text, "0x", "&h")
    If val(sTmp) > 0 Then
        oDrive.drvStartSector = CCur(val(sTmp)) / 10000@
    End If
    
'    If isTarget Then
        
        If (Trim(txtSector.Text) = "" Or Trim(txtSector.Text) = "0x") Then
            MsgBox LangLabel("WriteSector", "Please Write sector number in Decimal or Hex writing the prefix 0x"), vbOKOnly, LangLabel("Missing", "Missing Sector number")
        Else
'           oDrive.drvStartSector = CCur(val(sTmp)) / 10000@
            Unload Me
        End If
'    Else
'        Unload Me
'    End If
    
EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "cmdOKSD Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB

    
    
End Sub

Private Sub cmdRefresh_Click()

    FindDrives
    FillCombo

End Sub

Private Sub Form_Load()

On Error GoTo ERR_SUB

    m_bLoadingForm = True
    
    FindDrives
    FillCombo
    
    If cboDrives.ListCount > 0 Then
        cboDrives.ListIndex = 0
    End If
    
    'txtSector.Enabled = isTarget

    chkAllDrive.Value = IIf(g_bFindAllDrives, 1, 0)
    
    'txtSector.Enabled = True
    m_bLoadingForm = False
    
    
EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "SD FormLoad Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB
    
    
End Sub


Private Sub FillCombo()

    Dim oDrv As clsDrive

    lstDrives.Clear
    cboDrives.Clear

    If Not g_cDrives Is Nothing Then
        
        For Each oDrv In g_cDrives
            cboDrives.AddItem oDrv.drvName & "   " & oDrv.drvModel & " " & Format(CStr(oDrv.drvSize / (CLng(1024) * CLng(1024))), "#,### MByte ")
            oDrv.drvIndex = cboDrives.NewIndex
        Next
    
    End If
End Sub


'Procedures
Private Sub FillList()
    
    Dim prt As clsPartition

    lstDrives.Clear

    If Not oDrive Is Nothing Then
        
        For Each prt In oDrive.Partitions
            lstDrives.AddItem Trim(Split(Replace(prt.Name, "#", ""), ",")(1)) & " - " & Format(CStr((prt.Size * 10000) / (CLng(1024) * CLng(1024))), "#,### MByte ") & "(0x" & CStr(Hex(prt.StartingAddress / 512)) & ")"
        Next
    
    End If
    
End Sub

Private Sub lstDrives_Click()
 
    Dim tmp As String
    Dim val As Currency
    
    tmp = Split(lstDrives.Text, "(")(1)
    tmp = Left(tmp, Len(tmp) - 1)
    tmp = Replace(tmp, "0x", "&h")
    val = CCur(tmp)
    val = val + 32768@
    txtSector.Text = "0x" & Hex(val)
    
    
End Sub

