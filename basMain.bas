Attribute VB_Name = "basMain"
'****************************************************************************
'*
'*      EmuTool
'*
'* @Author: TheyKilledKenny
'* @Date:   05 Oct 2019
'*
'*  Funzioni implementate:
'*  Copia della partizione nascosta della emuMMC Atmosphere su file
'*      - il file emummc.ini deve puntare alla partizione interessata
'*          (cioè dal menu Kosmos deve essere selezionata la emunand che si vuole copiare)
'*      - i file sono utilizzabili come emuMMC su file
'*          - Struttura cartelle compatibile per drag&drop sulla root della sd
'*          - crea emummc.ini compatibile Atmosphere/Kosmos
'*          - crea file file_based compatibile con Kosmos
'*      - Indica qual'è la cartella "Nintendo" collegata alla partizione i copia
'*
'*  Copia della partizione nascosta della Emunand SXOS su file
'*      - Il file è utilizzabile come emunand su file
'*          - Struttura cartelle compatibile per drag&drop sulla root della sd
'*      - E' necessario copiare manualmente la cartella Emutendo per mantenere i software installati
'*
'*  - SX Check stato attuale Emunand. Se abilitata la emu su hidden, la emu su file non funzona
'*  - SX Inserito switch per cambiare tipo di emunand
'*
'*  - Ripristino partizione nascosta Atmosphere
'*    necessario indicare il settore di ripristino
'*      - Creazione emummc.ini adeguato alla partizione
'*
'*  - Ripristino partizione nascosta SxOs
'*
'*  Bachi noti:
'*  - necessario allineare i trunk per Atmosphere su File a 4MB, potrebbe essere sufficiente usare la corretta dimensione file
'*    al posto di usare quella di SX
'*
'*  Da implementare:
'*
'*  - possibilità di individuare le partizioni, il ciclo esiste già ed è il medesimo usato per individuare il nome lettera drive
'*
'*****************************************************************************************




'Simulate a static class in a more object oriented language
'This module is used as the main Drive list class
Public g_cDrives As Collection

'Public Const EMU_SX_1024_SECTOR As Currency = 0
'Public Const EMU_SX_1024_LENGHT As Long = 2

Public Const EMU_SX_ENABLE_SECTOR As Currency = 0.0001@

Public Const EMU_SX_BOOT0_SECTOR As Currency = 0.0002@
Public Const EMU_BOOT0_LENGHT As Long = 8192

'Public Const EMU_SX_BOOT1_SECTOR As Currency = 0.8194@
Public Const EMU_BOOT1_LENGHT As Long = 8192

'Public Const EMU_SX_NAND_SECTOR As Currency = 1.6386@
Public Const EMU_NAND_LENGHT As Long = 61071360

Public Const EMU_SX_NAND_TRUNK_SIZE As Long = 8388352    'Last trunk should be 2352896
Public Const EMU_SX_NAND_LAST_TRUNK_SIZE As Long = 2352896  'Last trunk should be 2352896

Public Const EMU_AMS_NAND_TRUNK_SIZE As Long = 8323072
Public Const EMU_AMS_NAND_LAST_TRUNK_SIZE As Long = 2809856

Public Const EMU_BLOCK_SIZE As Currency = 512

Public Const EMU_SX_BOOT0_PATH As String = "\sxos\Emunand\boot0.bin"
Public Const EMU_SX_BOOT1_PATH As String = "\sxos\Emunand\boot1.bin"
Public Const EMU_SX_RAWNA_PATH As String = "\sxos\Emunand\full.0#.bin"

Public Const EMU_ATM_BOOT0_PATH As String = "\eMMC\BOOT0"
Public Const EMU_ATM_BOOT1_PATH As String = "\eMMC\BOOT1"
Public Const EMU_ATM_RAWNA_PATH As String = "\eMMC\0#"
            
Public Enum enEmuType
    enNone = 0
    enAtmHidden = 1
    enAtmFile = 2
    enSxHidden = 3
    enSxFile = 4
    enHekBack = 5
End Enum
            
Public g_bFindAllDrives As Boolean
Public g_bStopOperations As Boolean




Public lblStatus    As Label


Sub Main()

    On Error GoTo ERR_SUB

    'Do Not find Hard Disk at the beginning
    'If true it display also hard disks wich is dangerous
    
    g_bFindAllDrives = False


    'Turno off stop flag
    g_bStopOperations = False
    

    'FindDrives

    frmMain.Show
    
    
EXIT_SUB:

    Exit Sub
    
ERR_SUB:

    Resume EXIT_SUB
    
    
    
End Sub


'Fill Drives collection with system retrieved informations
'TODO: Error checking
Public Sub FindDrives()

On Error GoTo ERR_SUB
    
    Dim oDrive As clsDrive
    Dim List
    Dim Object
    
    'Initialize the Drives collections
    Set g_cDrives = New Collection
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_DiskDrive")
    
    For Each Object In List
        If (Trim(LCase(Object.MediaType)) <> "fixed hard disk media") Or g_bFindAllDrives Then
            Set oDrive = New clsDrive
            oDrive.init Object
            g_cDrives.Add oDrive, oDrive.drvDeviceID
            
        End If
    
    Next

EXIT_SUB:

    Exit Sub
    
ERR_SUB:

    Resume EXIT_SUB

End Sub

'Ready for a multilanguage version
'All text should be passed by this function.
Public Function LangLabel(ByRef sKey As String, ByRef sDefault As String) As String
    LangLabel = sDefault
End Function

