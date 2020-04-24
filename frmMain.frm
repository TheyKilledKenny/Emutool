VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EmuTool"
   ClientHeight    =   5190
   ClientLeft      =   2055
   ClientTop       =   1785
   ClientWidth     =   12675
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   12675
   StartUpPosition =   2  'CenterScreen
   Tag             =   "AppTitle"
   Begin VB.Frame fraSxEmu 
      Caption         =   "SX OS:"
      Height          =   765
      Left            =   30
      TabIndex        =   21
      Top             =   3930
      Visible         =   0   'False
      Width           =   6225
      Begin VB.OptionButton optSetSx 
         Caption         =   "Disable Partition Emu"
         Height          =   435
         Index           =   1
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   23
         Tag             =   "DisableSx"
         Top             =   210
         Width           =   1785
      End
      Begin VB.OptionButton optSetSx 
         Caption         =   "Enable Partition Emu"
         Height          =   435
         Index           =   0
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   22
         Tag             =   "EnableSx"
         Top             =   210
         Width           =   1785
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Target:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Index           =   1
      Left            =   6360
      TabIndex        =   6
      Tag             =   "SelectDest"
      Top             =   60
      Width           =   6225
      Begin VB.TextBox txtSrcSd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmMain.frx":57E2
         Top             =   2760
         Width           =   5715
      End
      Begin VB.OptionButton optDest 
         Caption         =   "Atmosphere hidden Partition"
         Height          =   375
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   930
         Width           =   2565
      End
      Begin VB.OptionButton optDest 
         Caption         =   "Atmosphere File"
         Height          =   375
         Index           =   2
         Left            =   300
         TabIndex        =   9
         Top             =   1620
         Width           =   1755
      End
      Begin VB.OptionButton optDest 
         Caption         =   "SXOS hidden partition"
         Height          =   375
         Index           =   3
         Left            =   3780
         TabIndex        =   8
         Top             =   930
         Width           =   2115
      End
      Begin VB.OptionButton optDest 
         Caption         =   "SXOS File"
         Height          =   375
         Index           =   4
         Left            =   3780
         TabIndex        =   7
         Top             =   1620
         Width           =   1635
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000A&
         X1              =   150
         X2              =   6000
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Select Emunand or emuMMC Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Width           =   5715
      End
   End
   Begin VB.Frame fraStaturBar 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4860
      Width           =   12645
      Begin VB.Label lblStatusBar 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   12645
      End
      Begin VB.Label lblStatusBar 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "by TheyKilledKenny"
         Height          =   285
         Index           =   3
         Left            =   6360
         TabIndex        =   15
         Top             =   0
         Width           =   6285
      End
      Begin VB.Label lblStatusBar 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Current: RAWNAND07"
         Height          =   285
         Index           =   2
         Left            =   4230
         TabIndex        =   14
         Top             =   0
         Width           =   2145
      End
      Begin VB.Label lblStatusBar 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Blocksi:"
         Height          =   285
         Index           =   1
         Left            =   2220
         TabIndex        =   13
         Tag             =   "TotalBlock"
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label lblStatusBar 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blocchi Scritti: 65.535 "
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   2205
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10530
      TabIndex        =   17
      Tag             =   "Stop"
      Top             =   4230
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Source:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Index           =   0
      Left            =   30
      TabIndex        =   1
      Tag             =   "SelectSource"
      Top             =   60
      Width           =   6225
      Begin VB.TextBox txtSrcSd 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "frmMain.frx":57FC
         Top             =   2760
         Width           =   5715
      End
      Begin VB.OptionButton optSrc 
         Caption         =   "Hekate Backup Files"
         Height          =   375
         Index           =   5
         Left            =   300
         TabIndex        =   18
         Top             =   2130
         Width           =   3705
      End
      Begin VB.OptionButton optSrc 
         Caption         =   "SXOS File"
         Height          =   375
         Index           =   4
         Left            =   3660
         TabIndex        =   5
         Top             =   1620
         Width           =   1755
      End
      Begin VB.OptionButton optSrc 
         Caption         =   "SXOS hidden partition"
         Height          =   375
         Index           =   3
         Left            =   3660
         TabIndex        =   4
         Top             =   930
         Width           =   2325
      End
      Begin VB.OptionButton optSrc 
         Caption         =   "Atmosphere File"
         Height          =   375
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   1620
         Width           =   1875
      End
      Begin VB.OptionButton optSrc 
         Caption         =   "Atmosphere hidden Partition"
         Height          =   375
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   930
         Width           =   2625
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   180
         X2              =   6030
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Select Emunand or emuMMC Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   300
         Width           =   5715
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   0
      Tag             =   "Start"
      Top             =   4230
      Width           =   4005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************
' Scritto da TheyKilledKenny il 25 Sept 2019
'   Spostato procedure di DirectReadDriveNT e DirectWriteDriveNT nel modulo diskio e rese pubbliche
'   Completamente rivista la procedura di lettura dei dati
'
'*****************************************************************************************************
Option Explicit

'Frame Index
Private Const FRAME_SOURCE = 0
Private Const FRAME_DESTINATION = 1
Private Const FRAME_SD = 2
Private Const FRAME_FILE = 3

Private m_enSrcType     As enEmuType    'Contiene la scelta della sorgente
Private m_enDstType     As enEmuType    'Contiene la scelta della destinazione
Private m_oSourceSd   As clsDrive     'Contiene il drive attualmente selezionato
Private m_oTargetSd   As clsDrive     'Contiene il drive attualmente selezionato

Private m_sSourcePath   As String       'Contiene il percorso di destinazione dei file
Private m_sSourceNin    As String       'Contiene il percorso di destinazione dei file
Private m_sTargetNin    As String       'Contiene il percorso di destinazione dei file

Private m_sOutPath  As String       'Contiene il percorso di dove dover copiare i file finali

Private m_sBoot0Path    As String       'contiene il nome e percorso di boot0
Private m_sBoot1Path    As String       'contiene il nome e percorso di boot1
Private m_sRawNandPath  As String       'contiene il nome e percorso dei pezzi di rawnand

Private m_sSBoot0Path    As String       'contiene il nome e percorso di boot0
Private m_sSBoot1Path    As String       'contiene il nome e percorso di boot1
Private m_sSRawNandPath  As String       'contiene il nome e percorso dei pezzi di rawnand


Private m_oDrive        As clsDrive
Private m_bLoadingForm  As Boolean      'Is the form still loading? (or no operation on the control's events)


'Set the flag to stop the operation
Private Sub cmdStop_Click()
    
    If MsgBox(LangLabel("SureCancelOp", "Are you sure you want to stop the job?"), vbQuestion + vbYesNo, LangLabel("Sure", "Really sure?")) = vbYes Then
        g_bStopOperations = True
    Else
        g_bStopOperations = False
    End If
        
End Sub

'Catch the click on Target Type
Private Sub optDest_Click(Index As Integer)

    'riaccendi le options
    'EnableOpt

    m_enDstType = Index
    
    'Ripristina il drive da utilizzare
    txtSrcSd(1).Text = "Click to select SD Card"
    txtSrcSd(1).FontBold = True
    txtSrcSd(1).Alignment = vbCenter   'vbAlignLeft
    
End Sub

'Cattura l'evento di click sui pulsanti abilita e disabilita per SXOS
Private Sub optSetSx_Click(Index As Integer)

On Error GoTo ERR_SUB

    If Not m_bLoadingForm Then
    
        If Index = 0 Then
            m_oSourceSd.sxPartitionEnabled = True
        
        Else
            m_oSourceSd.sxPartitionEnabled = False
        End If
        
    End If
        
EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "optSetSX Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB
        
    
End Sub

'Intercetta il click sul tipo di Sorgente
' Abilita e disabilita i controlli sulla form in base alla scelta
'
Private Sub optSrc_Click(Index As Integer)

    'riaccendi le options
    EnableOpt
    
    'Ripristina il drive da utilizzare
    txtSrcSd(0).Text = "Click to select SD Card"
    txtSrcSd(0).FontBold = True
    txtSrcSd(0).Alignment = vbCenter   'vbAlignLeft
    
    Set m_oSourceSd = Nothing


    Select Case Index
    
        Case enAtmHidden
            'Spegni le destinazioni che non servono
            optDest(enAtmHidden).Value = False
            optDest(enAtmHidden).Enabled = False
            optDest(enSxHidden).Value = False
            optDest(enSxHidden).Enabled = False
            txtSrcSd(0).Text = "Click to select SD Card"
            m_enSrcType = enAtmHidden

        Case enAtmFile
            'Spegni le destinazioni che non servono
            optDest(enAtmFile).Value = False
            optDest(enAtmFile).Enabled = False
            txtSrcSd(0).Text = "Click to select Folder"
            m_enSrcType = enAtmFile


        Case enSxHidden
            'Spegni le destinazioni che non servono
            optDest(enAtmHidden).Value = False
            optDest(enAtmHidden).Enabled = False
            optDest(enSxHidden).Value = False
            optDest(enSxHidden).Enabled = False
            txtSrcSd(0).Text = "Click to select SD Card"
            m_enSrcType = enSxHidden
            
            
        Case enSxFile
            'Spegni le destinazioni che non servono
            optDest(enSxFile).Value = False
            optDest(enSxFile).Enabled = False
            txtSrcSd(0).Text = "Click to select Folder"
            m_enSrcType = enSxFile
            
            
        Case enHekBack
        
            m_enSrcType = enHekBack
            
        
    End Select
    
End Sub


Private Sub Form_Load()

On Error GoTo ERR_SUB

'Set starting flag to avoid other controls events
m_bLoadingForm = True
    
    m_enSrcType = enNone    'Reset selected src type
    m_enDstType = enNone    'Reset selected dst type
        
    'optSrc(enAtmHidden).Value = True
    EnableForm True

    Me.Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision

    SetInitialStatusBar

m_bLoadingForm = False


EXIT_SUB:

    Exit Sub
    
ERR_SUB:

    MsgBox "FormLoad Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB

End Sub


Private Sub cmdStart_Click()
    
On Error GoTo ERR_SUB
    
    EnableForm False
    g_bStopOperations = False
    
    
    If MsgBox("Please check data again!" & vbCrLf & "Are you sure you want to continue?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
    
        If (m_enSrcType = enAtmHidden Or m_enSrcType = enSxHidden) And _
            (m_enDstType = enAtmFile Or m_enDstType = enSxFile) Then
            PartitionToFile
        
        ElseIf (m_enSrcType = enAtmFile Or m_enSrcType = enSxFile Or m_enSrcType = enHekBack) And _
               (m_enDstType = enAtmHidden Or m_enDstType = enSxHidden) Then
            
            FileToPartition
        
        Else
            FileToFile
        
        End If
    End If
    
    
    
    
    
    EnableForm True


EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "cmdStartClick Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB



End Sub

'Enable/disable opt controls
Private Sub EnableOpt(Optional ByVal bEnable As Boolean = True, Optional iWich As Integer = 3)

    If (iWich And 2) > 0 Then
        optDest(enAtmHidden).Enabled = bEnable
        optDest(enAtmFile).Enabled = bEnable
        optDest(enSxHidden).Enabled = bEnable
        optDest(enSxFile).Enabled = bEnable
    End If
    If (iWich And 1) > 0 Then
        optSrc(enAtmHidden).Enabled = bEnable
        optSrc(enAtmFile).Enabled = bEnable
        optSrc(enSxHidden).Enabled = bEnable
        optSrc(enSxFile).Enabled = bEnable
    End If

End Sub

Private Sub EnableForm(ByVal bEnable As Boolean)
    
    cmdStart.Enabled = bEnable
    cmdStop.Enabled = Not bEnable
    
    'lstDrives.Enabled = bEnable
    Frame1(0).Enabled = bEnable
    Frame1(1).Enabled = bEnable

End Sub


Public Sub SetInitialStatusBar()
   
    fraStaturBar.Top = Me.ScaleHeight - fraStaturBar.Height
    
    lblStatusBar(4).Visible = True
    lblStatusBar(4).Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    lblStatusBar(4).Caption = lblStatusBar(4).Caption & " - " & LangLabel("StatusBarText", App.FileDescription & " by " & App.CompanyName & " (..." & App.Comments & ")")

End Sub

'Inizializza la status bar con il testo predefinito per il conteggio e ritorna il panel in cui scrivere il conto
Public Function SetCounterStatusBar(ByVal TotalSectors As Currency) As Label
    
    lblStatusBar(4).Visible = False
    lblStatusBar(0).Caption = "Written blocks: 0"
    lblStatusBar(1).Caption = "Total Block: " & CStr(TotalSectors)
    lblStatusBar(2).Caption = ""
    lblStatusBar(3).Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision

    Set SetCounterStatusBar = lblStatusBar(0)
    
End Function

'Set path and filenames according to user choice
Private Function SetProperPath() As Boolean


On Error GoTo ERR_SUB
    
    Dim tmpFileNo As Long
    
    If m_oSourceSd Is Nothing Then
        If Trim(m_sSourcePath) = "" Then
            MsgBox LangLabel("MissingSource", "Please choose a Source"), vbOKOnly, LangLabel("Errore", "Error")
            SetProperPath = False
            Exit Function
        End If
    End If


'Set source path
    If Trim(m_sSourcePath) = "" Then
        m_sSourcePath = m_oSourceSd.drvName & "\" & m_oSourceSd.drvEmuFolder
    End If
    
        Select Case m_enSrcType
        
            Case enAtmFile
                m_sSBoot0Path = m_sSourcePath & "\BOOT0" 'EMU_ATM_BOOT0_PATH
                m_sSBoot1Path = m_sSourcePath & "\BOOT1" 'EMU_ATM_BOOT1_PATH
                m_sSRawNandPath = m_sSourcePath & "\#" 'EMU_ATM_RAWNA_PATH
            
            Case enSxFile
                m_sSBoot0Path = m_sSourcePath & "\boot0.bin"    'EMU_SX_BOOT0_PATH
                m_sSBoot1Path = m_sSourcePath & "\boot1.bin"    'EMU_SX_BOOT1_PATH
                'm_sSRawNandPath = m_sSourcePath & "\full.0#.bin"    'EMU_SX_RAWNA_PATH
                m_sSRawNandPath = m_sSourcePath & "\full.#.bin"    'EMU_SX_RAWNA_PATH
                
            Case enHekBack
                m_sSBoot0Path = m_sSourcePath & "\BOOT0"
                m_sSBoot1Path = m_sSourcePath & "\BOOT1"
                m_sSRawNandPath = m_sSourcePath & "\rawnand.bin#"
            

            Case enAtmHidden
                m_sSBoot0Path = ""
                m_sSBoot1Path = ""
                m_sSRawNandPath = ""
            
            Case enSxHidden
                m_sSBoot0Path = ""
                m_sSBoot1Path = ""
                m_sSRawNandPath = ""
    
        End Select


    
'Set target path
    
    If m_oTargetSd Is Nothing Then
        If Trim(m_sOutPath) = "" Then
            MsgBox LangLabel("MissingTarget", "Please choose Target"), vbOKOnly, LangLabel("Errore", "Error")
            SetProperPath = False
            Exit Function
        End If
    End If
    
    If Trim(m_sOutPath) = "" Then
        m_sOutPath = m_oTargetSd.drvName & "\" & m_oTargetSd.drvEmuFolder
    End If
    
        Select Case m_enDstType
        
            Case enAtmFile
                m_sBoot0Path = m_sOutPath & "\emummc\HPE0" & EMU_ATM_BOOT0_PATH
                m_sBoot1Path = m_sOutPath & "\emummc\HPE0" & EMU_ATM_BOOT1_PATH
                m_sRawNandPath = m_sOutPath & "\emummc\HPE0" & EMU_ATM_RAWNA_PATH
            
                On Error Resume Next
                MkDir m_sOutPath & "\emummc"
                
                'Creo il file emummc.ini
                tmpFileNo = FreeFile
                Open m_sOutPath & "\emuMMC\emummc.ini" For Output As tmpFileNo
                Print #tmpFileNo, "[emummc]"
                Print #tmpFileNo, "enabled=1"
                Print #tmpFileNo, "sector=0x0"
                Print #tmpFileNo, "path=emuMMC/HPE0"
                Print #tmpFileNo, "id=0x0000"
                Print #tmpFileNo, "nintendo_path=emuMMC/HPE0/Nintendo"
                Print #tmpFileNo, ""
                Close tmpFileNo
                
                MkDir m_sOutPath & "\emummc\HPE0"
                Kill m_sOutPath & "\emummc\HPE0\*.*"
            
                tmpFileNo = FreeFile
                Open m_sOutPath & "\emuMMC\HPE0\file_based" For Output As tmpFileNo
                Close tmpFileNo
                
                MkDir m_sOutPath & "\emummc\HPE0\eMMC"
                Kill m_sOutPath & "\emummc\HPE0\eMMC\*.*"
                On Error GoTo 0
            
            Case enSxFile
                m_sBoot0Path = m_sOutPath & EMU_SX_BOOT0_PATH
                m_sBoot1Path = m_sOutPath & EMU_SX_BOOT1_PATH
                m_sRawNandPath = m_sOutPath & EMU_SX_RAWNA_PATH
                
                'Elimina eventuali precedenti file
                On Error Resume Next
                MkDir m_sOutPath & "\sxos"
                Kill m_sOutPath & "\sxos\*.*"
                MkDir m_sOutPath & "\sxos\Emunand"
                Kill m_sOutPath & "\sxos\Emunand\*.*"
                On Error GoTo 0
            
            Case enAtmHidden
                m_sBoot0Path = ""
                m_sBoot1Path = ""
                m_sRawNandPath = ""
                
            Case enSxHidden
                m_sBoot0Path = ""
                m_sBoot1Path = ""
                m_sRawNandPath = ""
    
        End Select
    
    
    'Remove double slash to avoid compatibility issues
    m_sBoot0Path = Replace(m_sBoot0Path, "\\", "\")
    m_sBoot1Path = Replace(m_sBoot1Path, "\\", "\")
    m_sRawNandPath = Replace(m_sRawNandPath, "\\", "\")
    
    m_sSBoot0Path = Replace(m_sSBoot0Path, "\\", "\")
    m_sSBoot1Path = Replace(m_sSBoot1Path, "\\", "\")
    m_sSRawNandPath = Replace(m_sSRawNandPath, "\\", "\")
    
    
    SetProperPath = True
        
EXIT_SUB:
    Exit Function
    
ERR_SUB:
    MsgBox "SetPropetPath Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB

        
End Function


'Read data from a starting sector in the selected drive and write it in files
Private Sub PartitionToFile()

    Dim cini                As cInifile       'Classe di gestione dei file INI
    'Dim oDrv               As clsDrive
    
    Dim tmpFileNo           As Long
    Dim iTrunkNum           As Integer
    Dim startSector         As Currency
    Dim hbfFile             As HugeBinaryFile
    Dim tmr                 As Single
    Dim eTmr                As Single
    Dim rawnandTrunkSize    As Currency
    
    On Error GoTo ERR_SUB

    tmr = Timer
    
    'If no drive selected, info the user and skip all.
    If m_oSourceSd Is Nothing Then
        MsgBox LangLabel("MsgSelectDrive", "Please select Drive"), vbOKCancel + vbInformation, LangLabel("Error", "Error")
        Exit Sub
    ElseIf SetProperPath = False Then
        'Prosegue solo se è possibile impostare i percorsi corretti
        'MsgBox LangLabel("ERRPath", "Something went wrong in setting and creting Folders"), vbOKCancel + vbInformation, LangLabel("Error", "Error")
        Exit Sub
    Else
    
        startSector = m_oSourceSd.drvStartSector

        
        '------------------------------------------------------------------------------------ boot0.bin
        'Inizializzo i pannelli della status bar con il valore di settori
        Set lblStatus = SetCounterStatusBar(EMU_BOOT0_LENGHT)
        lblStatusBar(2).Caption = " Current: BOOT0 "
        
        tmpFileNo = FreeFile
        Open m_sBoot0Path For Binary As tmpFileNo Len = 1
    
        m_oSourceSd.readDataToFileId startSector, EMU_BOOT0_LENGHT, tmpFileNo
    
        Close tmpFileNo
        
        If g_bStopOperations Then Exit Sub
        
        '------------------------------------------------------------------------------------ boot1.bin
        'Inizializzo i pannelli della status bar con il valore di settori
        lblStatusBar(1).Caption = "Total Blocks: " & CStr(EMU_BOOT1_LENGHT)
        lblStatusBar(2).Caption = " Current: BOOT1 "
        
        tmpFileNo = FreeFile
        Open m_sBoot1Path For Binary As tmpFileNo Len = 1
        startSector = startSector + (CCur(EMU_BOOT0_LENGHT) / 10000@)
        
        m_oSourceSd.readDataToFileId startSector, EMU_BOOT1_LENGHT, tmpFileNo
    
        Close tmpFileNo
        
        If g_bStopOperations Then Exit Sub

        
        '----------------------------------------------------------------------------------- rawnand.bin
        startSector = startSector + (CCur(EMU_BOOT1_LENGHT) / 10000@)
        
        If m_enDstType = enAtmFile Then
            rawnandTrunkSize = EMU_AMS_NAND_TRUNK_SIZE
        Else
            rawnandTrunkSize = EMU_SX_NAND_TRUNK_SIZE
        End If
        
        
        On Error GoTo ERR_FILE
        For iTrunkNum = 0 To 6
        
            Set hbfFile = New HugeBinaryFile
            hbfFile.OpenFile Replace(m_sRawNandPath, "#", CStr(iTrunkNum))
    
            'Inizializzo i pannelli della status bar con il valore di settori
            lblStatusBar(1).Caption = "Total Blocks: " & Format(CStr(rawnandTrunkSize), "#,###")
            lblStatusBar(2).Caption = " Current Rawnand: 0" & CStr(iTrunkNum) & "/07"
            
            m_oSourceSd.readDataToFileId startSector, rawnandTrunkSize, -1, hbfFile 'tmpFileNo
            startSector = startSector + (CCur(rawnandTrunkSize) / 10000@)
    
            hbfFile.Flush
            hbfFile.CloseFile
        
            If g_bStopOperations Then Exit Sub

    
        Next
        On Error GoTo 0

        If m_enDstType = enAtmFile Then
            rawnandTrunkSize = EMU_AMS_NAND_LAST_TRUNK_SIZE
        Else
            rawnandTrunkSize = EMU_SX_NAND_LAST_TRUNK_SIZE
        End If


        'Last trunk
        'Inizializzo i pannelli della status bar con il valore di settori
        lblStatusBar(0).Caption = "Total Block: " & CStr(rawnandTrunkSize)
        lblStatusBar(2).Caption = " Current Rawnand: 07/07"
        
        Set hbfFile = New HugeBinaryFile
        hbfFile.OpenFile Replace(m_sRawNandPath, "#", "7")
    
        m_oSourceSd.readDataToFileId startSector, rawnandTrunkSize, -1, hbfFile
        
        hbfFile.Flush
        hbfFile.CloseFile
               
        '----------------------------------------------------------------------------------- rawnand.bin end
        
        
        eTmr = Timer
    
        If m_enDstType = enAtmFile Then
        
            MsgBox LangLabel("NinFoldWarn", "Warning!" & vbCrLf & _
                               "Remember to copy " & m_oSourceSd.drvNinFolder & " Folder" & vbCrLf & _
                               "From SD Card" & vbCrLf & _
                               "To " & m_sOutPath & "emuMMC\HPE0\Nintendo"), vbOKOnly, "Copy Folder"
        
        Else
        
            MsgBox LangLabel("NinFoldWarn", "Warning!" & vbCrLf & _
                               "Remember to copy Emutendo Folder from SD Card" & vbCrLf & _
                               "To " & m_sOutPath & "Emutendo"), vbOKOnly, "Copy Folder" ', vbOKOnly, "Copy Folder"
        
        End If
    
    End If


EXIT_SUB:

    SetInitialStatusBar
    MsgBox LangLabel("EndDump", "Hey, I did it in") & " " & _
    CStr((eTmr - tmr) \ 60) & "min. " & CStr(CInt((eTmr - tmr) - (((eTmr - tmr) \ 60) * 60))) & "sec." _
    , vbOKOnly, LangLabel("EndDumpT", "Finish")
    
    
    Exit Sub
    
ERR_SUB:

    If Err.Number = 13 Then
        MsgBox LangLabel("BadIni", "emummc.ini file is not properly configured for hidden partition emu"), vbOKOnly + vbExclamation, "Error!"
        Resume EXIT_SUB
    Else
        MsgBox LangLabel("Errore", "Errore") & ": " & Err.Number & " - " & Err.Description
        Resume EXIT_SUB
    End If


ERR_FILE:
    MsgBox Err.Number & " - " & Err.Description
    Resume EXIT_SUB


End Sub



'Procedure to byte copy from file to a partition
Private Sub FileToPartition()
    
    Dim hbfSrc          As HugeBinaryFile    'File Sorgente
    Dim startSector     As Currency
    Dim SrcTrunkTotal   As Currency
    Dim idx             As Integer
    Dim sPoint          As String
    Dim SrcTrunkNum     As Integer
    Dim sBuffer(0 To 3)  As Byte
    Dim tmr             As Single
    Dim eTmr            As Single
    Dim sTmp            As String
    Dim tmpFileNo       As Long
    Dim bSingleFile   As Boolean
    
On Error GoTo ERR_SUB
    
    tmr = Timer
    
    'If no drive selected, info the user and skip all.
    If m_oTargetSd Is Nothing Then
        MsgBox LangLabel("MsgSelectDrive", "Please select Target Drive"), vbOKCancel + vbInformation, LangLabel("Error", "Error")
        Exit Sub
    ElseIf SetProperPath = False Then
        'Prosegue solo se è possibile impostare i percorsi corretti
        'MsgBox LangLabel("ERRPath", "Something went wrong in setting and creting Folders"), vbOKCancel + vbInformation, LangLabel("Error", "Error")
        Exit Sub
    Else
    
        startSector = m_oTargetSd.drvStartSector
        
'''''''''------------------------------------------------------------------------------------ boot0.bin
        'Set Initial file from user choice
        Set hbfSrc = New HugeBinaryFile
        Set lblStatus = SetCounterStatusBar(EMU_BOOT0_LENGHT)
        lblStatusBar(2).Caption = " Current: BOOT0 "
        
        If hbfSrc.OpenFile(m_sSBoot0Path, True) Then
            SrcTrunkTotal = hbfSrc.FileLen

             m_oTargetSd.readDataFromFileId startSector, hbfSrc
            
            hbfSrc.CloseFile
            Set hbfSrc = Nothing
        Else
            MsgBox "BOOT0 not found!", vbExclamation + vbOKOnly, "File not Found"
            SetInitialStatusBar
            Exit Sub
        End If
                      
        'startSector = startSector + (CCur(EMU_BOOT0_LENGHT) / 10000@)
        startSector = startSector + (SrcTrunkTotal / EMU_BLOCK_SIZE) / 10000@
           
'''''''''------------------------------------------------------------------------------------ boot1.bin
        'Set Initial file from user choice
        Set hbfSrc = New HugeBinaryFile
        lblStatusBar(1).Caption = "Total Blocks: " & Format(CStr(EMU_BOOT1_LENGHT), "#,###")
        lblStatusBar(2).Caption = " Current: BOOT1 "

        If hbfSrc.OpenFile(m_sSBoot1Path, True) Then
            SrcTrunkTotal = hbfSrc.FileLen
            
            m_oTargetSd.readDataFromFileId startSector, hbfSrc
            
            hbfSrc.CloseFile
            Set hbfSrc = Nothing
        Else
            MsgBox "BOOT1 not found!", vbExclamation + vbOKOnly, "File not Found"
            SetInitialStatusBar
            Exit Sub
        End If
                      
        'startSector = startSector + (CCur(EMU_BOOT1_LENGHT) / 10000@)
        startSector = startSector + (SrcTrunkTotal / EMU_BLOCK_SIZE) / 10000@
              



        '----------------------------------------------------------------------------------- rawnand.bin
'Avvio una copia di tutti i byte possibili
'Sistemo i puntatori
'se puntatore sorgente è arrivato alla fine allora cambio file sorgente
'se puntatore destinatario è arrivato allora cambio destinatario

        SrcTrunkTotal = 0
        'DstTrunkNum = -1
        
        'Controllo se esiste un file unico di backup hekate
        If m_enSrcType = enHekBack Then
            On Error GoTo HEKCHECKERR
            tmpFileNo = FreeFile
            bSingleFile = True
            Open Replace(m_sSRawNandPath, "#", "") For Input As tmpFileNo
            Close #tmpFileNo
            On Error GoTo ERR_SUB
        End If
        
        'Se è un backup Hekate con singolo file allora prendi solo quello
        If m_enSrcType = enHekBack And bSingleFile Then
        
            Set hbfSrc = New HugeBinaryFile
            If hbfSrc.OpenFile(Replace(m_sSRawNandPath, "#", ""), True) Then
                SrcTrunkNum = SrcTrunkNum + 1
                SrcTrunkTotal = hbfSrc.FileLen
                
                'Inizializzo i pannelli della status bar con il valore di settori
                lblStatusBar(1).Caption = "Total Blocks: " & Format(CStr(SrcTrunkTotal / EMU_BLOCK_SIZE), "#,###")
                lblStatusBar(2).Caption = " Current Rawnand: 0" & CStr(SrcTrunkNum - 1) & "/01"

                'Avvio una copia dei byte che devo leggete
                m_oTargetSd.readDataFromFileId startSector, hbfSrc
                
                'startSector = startSector + (CCur(EMU_BOOT1_LENGHT) / 10000@)
                startSector = startSector + (SrcTrunkTotal / EMU_BLOCK_SIZE) / 10000@
                hbfSrc.CloseFile
            End If
    
        'altrimenti prendi i file splittati come al solito.
        Else
                    
            If m_enSrcType = enHekBack Then sPoint = "." Else sPoint = ""
        
            For idx = 0 To 50
                Set hbfSrc = New HugeBinaryFile
                'If hbfSrc.OpenFile(Replace(m_sSRawNandPath, "#", CStr(SrcTrunkNum)), True) Then
                If hbfSrc.OpenFile(Replace(m_sSRawNandPath, "#", sPoint & Format(CStr(SrcTrunkNum), "00")), True) Then
                
                    SrcTrunkNum = SrcTrunkNum + 1
                    SrcTrunkTotal = hbfSrc.FileLen
                    
                    'Inizializzo i pannelli della status bar con il valore di settori
                    lblStatusBar(1).Caption = "Total Blocks: " & Format(CStr(SrcTrunkTotal / EMU_BLOCK_SIZE), "#,###")
                    lblStatusBar(2).Caption = " Current Rawnand: 0" & CStr(SrcTrunkNum - 1) & "/07"
    
                    'Avvio una copia dei byte che devo leggete
                    m_oTargetSd.readDataFromFileId startSector, hbfSrc
                    
                    'startSector = startSector + (CCur(EMU_BOOT1_LENGHT) / 10000@)
                    startSector = startSector + (SrcTrunkTotal / EMU_BLOCK_SIZE) / 10000@
                    hbfSrc.CloseFile
                Else
                    If idx > 0 And idx < 8 Then
                        MsgBox "Last parsed file is: " & Replace(m_sSRawNandPath, "#", CStr(SrcTrunkNum - 1))
                    ElseIf idx = 0 Then
                        MsgBox "No File found!"
                    End If
                    Exit For
                End If
                
            Next
            
        End If
    
        eTmr = Timer
    
        If m_enDstType = enAtmHidden Then
        
                On Error Resume Next
                MkDir m_oTargetSd.drvName & "\emummc"
                
                'Creo il file emummc.ini
                tmpFileNo = FreeFile
                Open m_oTargetSd.drvName & "\emuMMC\emummc.ini" For Output As tmpFileNo
                Print #tmpFileNo, "[emummc]"
                Print #tmpFileNo, "enabled=1"
                Print #tmpFileNo, "sector=0x" & CStr(Hex(m_oTargetSd.drvStartSector * 10000))
                Print #tmpFileNo, "path=emuMMC/RAW9"
                Print #tmpFileNo, "id=0x0000"
                Print #tmpFileNo, "nintendo_path=emuMMC/RAW9/Nintendo"
                Print #tmpFileNo, ""
                Close tmpFileNo
                
                MkDir m_oTargetSd.drvName & "\emummc\RAW9"
                Kill m_oTargetSd.drvName & "\emummc\RAW9\*.*"
            
                'm_oTargetSd.drvStartSector = 70467.9936
            
                sTmp = CStr(Hex(m_oTargetSd.drvStartSector * 10000))
                                
                sBuffer(3) = val("&h" & Mid(sTmp, 1, 2))
                sBuffer(2) = val("&h" & Mid(sTmp, 3, 2))
                sBuffer(1) = val("&h" & Mid(sTmp, 5, 2))
                sBuffer(0) = val("&h" & Mid(sTmp, 7, 2))


'0x2a009000
'00 90 00 2A

'                Set hbfSrc = New HugeBinaryFile
'                hbfSrc.OpenFile m_oTargetSd.drvName & "\emuMMC\RAW9\raw_based", False
'                hbfSrc.WriteBytes sBuffer
'                hbfSrc.Flush
'                hbfSrc.CloseFile

                tmpFileNo = FreeFile
                Open m_oTargetSd.drvName & "\emuMMC\RAW9\raw_based" For Binary Access Write As #tmpFileNo
                Put #tmpFileNo, , sBuffer
                Close tmpFileNo

                On Error GoTo 0
        
            MsgBox LangLabel("NinFoldWarn", "Warning!" & vbCrLf & _
                               "Remember to copy Nintendo Folder to SD Card" & vbCrLf & _
                               "in " & m_oTargetSd.drvName & "\emuMMC\RAW9\Nintendo"), vbOKOnly, "Copy Folder"
        
        Else
        
            MsgBox LangLabel("NinFoldWarn", "Warning!" & vbCrLf & _
                               "remember to copy Emutendo Folder To SD Card"), vbOKOnly, "Copy Folder"
        
        End If
           
    End If
    
    SetInitialStatusBar
    
    MsgBox LangLabel("EndDump", "Hey, I did it in") & " " & CStr((eTmr - tmr) \ 60) & "min. " & CStr(CInt((eTmr - tmr) - (((eTmr - tmr) \ 60) * 60))) & "sec.", vbOKOnly, LangLabel("EndDumpT", "Finish")


EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB

HEKCHECKERR:
    bSingleFile = False
    Resume Next
    
End Sub


'Procedure to "convert" from one type of file to another.
Private Sub FileToFile()

    Dim hbfSrc          As HugeBinaryFile    'File Sorgente
    Dim hbfDst          As HugeBinaryFile    'File Destinatario
    Dim aBuffer()       As Byte              'Buffer di lettura/scrittura
    Dim sSrcPath        As String
    Dim SrcTrunkTotal   As Currency
    Dim DstTrunkPointer As Currency
    Dim DstTrunkTotal   As Currency
    Dim DstTrunkNum     As Integer
    Dim SrcTrunkPointer As Currency
    Dim SrcTrunkNum     As Integer
    
    Dim BytesWritten    As Currency
    Dim BytesToWrite    As Currency
    Dim rawnandTrunkSize As Currency

    Dim tmpFileNo       As Integer
    Dim tmr             As Single
    Dim idx             As Integer
    Dim sPoint          As String
    
    On Error GoTo ERR_SUB

    tmr = Timer

    'Prosegue solo se è possibile impostare i percorsi corretti
    If SetProperPath = True Then
        
'''''''''------------------------------------------------------------------------------------ boot0.bin
        'Set Initial file from user choice
        Set hbfSrc = New HugeBinaryFile
        Set lblStatus = SetCounterStatusBar(EMU_BOOT0_LENGHT)
        lblStatusBar(2).Caption = " Current: BOOT0 "

        
        
        If hbfSrc.OpenFile(m_sSBoot0Path, True) Then
            SrcTrunkTotal = hbfSrc.FileLen
            ReDim aBuffer(0 To SrcTrunkTotal - 1)
            hbfSrc.ReadBytes aBuffer
        
            Set hbfDst = New HugeBinaryFile
            hbfDst.OpenFile m_sBoot0Path, False
            hbfDst.WriteBytes aBuffer
            hbfDst.Flush
        Else
            MsgBox "BOOT0 not found!", vbExclamation + vbOKOnly, "File not Found"
            SetInitialStatusBar
            Exit Sub
        End If
           
           
           
'''''''''------------------------------------------------------------------------------------ boot1.bin
        'Inizializzo i pannelli della status bar con il valore di settori
        lblStatusBar(1).Caption = "Total Blocks: " & CStr(EMU_BOOT1_LENGHT)
        lblStatusBar(2).Caption = " Current: BOOT1 "
        
        'Set Initial file from user choice
        Set hbfSrc = New HugeBinaryFile
        If hbfSrc.OpenFile(m_sSBoot1Path, True) Then
            SrcTrunkTotal = hbfSrc.FileLen
            ReDim aBuffer(0 To SrcTrunkTotal - 1)
            hbfSrc.ReadBytes aBuffer
        
            Set hbfDst = New HugeBinaryFile
            hbfDst.OpenFile m_sBoot1Path, False
            hbfDst.WriteBytes aBuffer
            hbfDst.Flush
        Else
            MsgBox "BOOT1 not found!", vbExclamation + vbOKOnly, "File not Found"
            SetInitialStatusBar
            Exit Sub
        End If


        '----------------------------------------------------------------------------------- rawnand.bin

        'Inizializzo i pannelli della status bar con il valore di settori
        lblStatusBar(1).Caption = "Total Blocks: 7"  ' & CStr(7)
        sSrcPath = m_sSRawNandPath
        SrcTrunkTotal = 0
        
        DstTrunkNum = -1
        
        'Set the correct trunk size based on type of destination
        If m_enDstType = enAtmFile Then
            DstTrunkTotal = EMU_AMS_NAND_TRUNK_SIZE * EMU_BLOCK_SIZE
        Else
            DstTrunkTotal = EMU_SX_NAND_TRUNK_SIZE * EMU_BLOCK_SIZE
        End If
        
        
        If m_enSrcType = enHekBack Then sPoint = "." Else sPoint = ""
        
        
        DstTrunkPointer = DstTrunkTotal 'To start the initialization of the first file in the second if


'Determino il segmento minore
'Avvio una copia di tutti i byte possibili
'Sistemo i puntatori
'se puntatore sorgente è arrivato alla fine allora cambio file sorgente
'se puntatore destinatario è arrivato allora cambio destinatario


'If hbfSrc.OpenFile(Replace(m_sSRawNandPath, "#", sPoint & Format(CStr(SrcTrunkNum), "00")), True) Then


        Do
            If SrcTrunkPointer >= SrcTrunkTotal Then
                'Carico il file sorgente
                Set hbfSrc = New HugeBinaryFile
                On Error Resume Next
                If Not hbfSrc.OpenFile(Replace(sSrcPath, "#", sPoint & Format(CStr(SrcTrunkNum), "00")), True) Then
                    MsgBox "No more source file availlable" & vbCrLf & "Last parsed file is: " & Replace(sSrcPath, "#", sPoint & Format(CStr(SrcTrunkNum - 1), "00"))
                    Exit Do
                End If
                SrcTrunkNum = SrcTrunkNum + 1
                SrcTrunkTotal = hbfSrc.FileLen
                SrcTrunkPointer = 0
            End If
            
            If DstTrunkPointer >= DstTrunkTotal Then
                'Creo il riferimento al file destinatario
                DstTrunkNum = DstTrunkNum + 1
                If DstTrunkNum > 7 Then Exit Do
                
                If Not hbfDst Is Nothing Then hbfDst.CloseFile
                Set hbfDst = New HugeBinaryFile
                hbfDst.OpenFile Replace(m_sRawNandPath, "#", CStr(DstTrunkNum)), False
                DstTrunkPointer = 0
            End If
    
    
            ' Determina il minore tra cosa resta da leggere e da scrivere
            If (DstTrunkTotal - DstTrunkPointer) < (SrcTrunkTotal - SrcTrunkPointer) Then
                'meno da scrivere che da leggere
                BytesToWrite = (DstTrunkTotal - DstTrunkPointer)
            Else
                'meno da leggere che da scrivere
                BytesToWrite = (SrcTrunkTotal - SrcTrunkPointer)
            End If

            'Inizializzo i pannelli della status bar con il valore di settori
            lblStatusBar(1).Caption = "Total Blocks: " & Format(CStr(DstTrunkTotal / EMU_BLOCK_SIZE), "#,###")
            lblStatusBar(2).Caption = " Current Rawnand: 0" & CStr(DstTrunkNum) & "/07"
 
            
            'Avvio una copia dei byte che devo leggete
            BytesWritten = hbfSrc.readDataToFileId(SrcTrunkPointer, BytesToWrite, DstTrunkPointer, hbfDst)
            hbfDst.Flush
            
            SrcTrunkPointer = SrcTrunkPointer + BytesWritten
            DstTrunkPointer = DstTrunkPointer + BytesWritten
            
        Loop While DstTrunkNum < 8 'Next
        
        If Not hbfSrc Is Nothing Then hbfSrc.CloseFile
        If Not hbfDst Is Nothing Then
            hbfDst.Flush
            hbfDst.CloseFile
        End If
        
  

''''''''
''''''''        'Last trunk
''''''''        'Inizializzo i pannelli della status bar con il valore di settori
''''''''        lblStatusBar(0).Caption = "Total Block: " & CStr(EMU_SX_NAND_LAST_TRUNK_SIZE)
''''''''        lblStatusBar(2).Caption = " Current: Rawnand07"
''''''''
''''''''        Set hbfFile = New HugeBinaryFile
''''''''        hbfFile.OpenFile Replace(m_sRawNandPath, "#", "7")
''''''''
''''''''        m_oDrive.readDataToFileId startSector, EMU_SX_NAND_LAST_TRUNK_SIZE, -1, hbfFile
''''''''
''''''''        hbfFile.Flush
''''''''        hbfFile.CloseFile
''''''''
        '----------------------------------------------------------------------------------- rawnand.bin end




''''        If m_enDstType = enAtmFile Then
''''
''''            'Creo il file emummc.ini
''''            tmpFileNo = FreeFile
''''            Open m_sOutPath & "\emuMMC\emummc.ini" For Output As tmpFileNo
''''            Print #tmpFileNo, "[emummc]"
''''            Print #tmpFileNo, "enabled=1"
''''            Print #tmpFileNo, "sector=0x0"
''''            Print #tmpFileNo, "path=emuMMC/HPE0"
''''            Print #tmpFileNo, "id=0x0000"
''''            Print #tmpFileNo, "nintendo_path=emuMMC/HPE0/Nintendo"
''''            Print #tmpFileNo, ""
''''            Close tmpFileNo
''''
''''
''''            tmpFileNo = FreeFile
''''            Open App.Path & "\emuMMC\HPE0\file_based" For Output As tmpFileNo
''''            Close tmpFileNo
''''
''''        End If


    End If


EXIT_SUB:

    SetInitialStatusBar
    MsgBox LangLabel("EndDump", "Hey, I did it in") & " " & CStr(Timer - tmr) & "sec.", vbOKOnly, LangLabel("EndDumpT", "Finish")
    Exit Sub
    
ERR_SUB:

    If Err.Number = 13 Then
        MsgBox LangLabel("BadIni", "emummc.ini file is not properly configured for hidden partition emu"), vbOKOnly + vbExclamation, "Error!"
        Resume EXIT_SUB
    Else
        MsgBox LangLabel("Errore", "Errore") & ": " & Err.Number & " - " & Err.Description
        Resume Next ' EXIT_SUB
    End If


ERR_FILE:
    MsgBox Err.Number & " - " & Err.Description
    Resume EXIT_SUB


End Sub


Private Sub txtSrcSd_Click(Index As Integer)

On Error GoTo ERR_SUB

    'Cliccato su sorgente
    If Index = 0 Then
        Select Case m_enSrcType
            Case enAtmFile, enSxFile, enHekBack
                frmFolder.sEmuFolder = m_sSourcePath
                frmFolder.isTarget = False
                frmFolder.Show vbModal
                If Trim(frmFolder.sEmuFolder) = "" Then Exit Sub
                
                m_sSourcePath = frmFolder.sEmuFolder
                frmFolder.sEmuFolder = ""
                
                Set m_oSourceSd = Nothing
                txtSrcSd(Index).Text = "Path: " & vbCrLf & m_sSourcePath
                'txtSrcSd(Index).FontBold = False
                txtSrcSd(Index).Alignment = vbLeftJustify
                fraSxEmu.Visible = False
            
            Case enAtmHidden
                frmSD.isTarget = False
                frmSD.enTipo = enAtmHidden
                frmSD.Show vbModal
                If frmSD.oDrive Is Nothing Then Exit Sub
                Set m_oSourceSd = frmSD.oDrive
                Set frmSD.oDrive = Nothing
                
                txtSrcSd(Index).Text = "Drive: " & m_oSourceSd.drvName & " (" & m_oSourceSd.drvModel & ")"
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Partition Sector: 0x" & CStr(Hex(m_oSourceSd.drvStartSector * 10000@))
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Nintendo Path: " & CStr(m_oSourceSd.drvNinFolder)
                'txtSrcSd(Index).FontBold = False
                txtSrcSd(Index).Alignment = vbLeftJustify
    
                fraSxEmu.Visible = False
                m_sSourcePath = ""
    
            Case enSxHidden
                frmSD.isTarget = False
                frmSD.enTipo = enSxHidden
                frmSD.Show vbModal
                If frmSD.oDrive Is Nothing Then Exit Sub
                Set m_oSourceSd = frmSD.oDrive
                Set frmSD.oDrive = Nothing
                
                'Imposta il settore di default
                m_oSourceSd.drvStartSector = 0.0002
                m_oSourceSd.drvEmuFolder = "sxos\Emunand"
                m_oSourceSd.drvNinFolder = "Emutendo"
                txtSrcSd(Index).Text = "Drive: " & m_oSourceSd.drvName & " (" & m_oSourceSd.drvModel & ")"
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Partition Sector: 0x" & CStr(Hex(m_oSourceSd.drvStartSector * 10000))
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Nintendo Path: " & CStr(m_oSourceSd.drvNinFolder)
                'txtSrcSd(Index).FontBold = False
                txtSrcSd(Index).Alignment = vbLeftJustify
    
                fraSxEmu.Visible = True
                m_sSourcePath = ""
                
                If Not m_oSourceSd Is Nothing Then
                    m_bLoadingForm = True
                    If m_oSourceSd.sxPartitionEnabled Then
                        optSetSx(0).Value = True
                    Else
                        optSetSx(1).Value = True
                    End If
                    m_bLoadingForm = False
                End If
                
            Case Else
            
        End Select

    'Cliccato su destinatario
    Else
        
        Select Case m_enDstType
            Case enAtmFile, enSxFile
                frmFolder.sEmuFolder = m_sOutPath
                frmFolder.isTarget = True
                frmFolder.Show vbModal
                If Trim(frmFolder.sEmuFolder) = "" Then Exit Sub
                m_sOutPath = frmFolder.sEmuFolder
                frmFolder.sEmuFolder = ""
                
                Set m_oTargetSd = Nothing
                                
                txtSrcSd(Index).Text = "Path: " & vbCrLf & m_sOutPath
                'txtSrcSd(Index).FontBold = False
                txtSrcSd(Index).Alignment = vbLeftJustify
                
            Case enAtmHidden
                frmSD.isTarget = True
                frmSD.enTipo = enAtmHidden
                frmSD.Show vbModal
                If frmSD.oDrive Is Nothing Then Exit Sub
                Set m_oTargetSd = frmSD.oDrive
                
                'm_oTargetSd.drvStartSector = m_oTargetSd.drvStartSector + (32768 / 10000@)
                
                Set frmSD.oDrive = Nothing
                m_sOutPath = ""
                
                txtSrcSd(Index).Text = "Drive: " & m_oTargetSd.drvName & " (" & m_oTargetSd.drvModel & ")"
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Partition Sector: 0x" & CStr(Hex(m_oTargetSd.drvStartSector * 10000@))
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Nintendo Path: " & CStr(m_oTargetSd.drvNinFolder)
                'txtSrcSd(Index).FontBold = False
                txtSrcSd(Index).Alignment = vbLeftJustify
    
                m_sOutPath = ""
            
            Case enSxHidden
                frmSD.isTarget = True
                frmSD.enTipo = enSxHidden
                frmSD.Show vbModal
                If frmSD.oDrive Is Nothing Then Exit Sub
                Set m_oTargetSd = frmSD.oDrive
                Set frmSD.oDrive = Nothing
                m_sOutPath = ""
                
                'Imposta il settore di default
                m_oTargetSd.drvStartSector = 0.0002
                m_oTargetSd.drvEmuFolder = "sxos\Emunand"
                m_oTargetSd.drvNinFolder = "Emutendo"
                txtSrcSd(Index).Text = "Drive: " & m_oTargetSd.drvName & " (" & m_oTargetSd.drvModel & ")"
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Partition Sector: 0x" & CStr(Hex(m_oTargetSd.drvStartSector * 10000@))
                txtSrcSd(Index).Text = txtSrcSd(Index).Text & vbCrLf & "Nintendo Path: " & CStr(m_oTargetSd.drvNinFolder)
                'txtSrcSd(Index).FontBold = False
                txtSrcSd(Index).Alignment = vbLeftJustify
                m_sOutPath = ""
                
            Case Else
            
        End Select

    End If
    
    'm_oSelectedSd
EXIT_SUB:
    Exit Sub
    
ERR_SUB:
    MsgBox "txtSrcSD Unexpected Error n." & Err.Number & vbCrLf & Err.Description
    Resume EXIT_SUB


End Sub
