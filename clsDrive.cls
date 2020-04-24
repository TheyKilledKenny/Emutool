VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*
'* Class Created by TheyKilledKenny on 28 Sept 2019
'*
'* heavily modified to correct the 64bit sector number for big SD
'* Drive information and Read/write procedures
'*
'*****************************************************

'* Derived from the original work of:
'
'*****************************************************************
' Module for performing Direct Read/Write access to disk sectors
'
' Written by Arkadiy Olovyannikov (ark@fesma.ru)
'*****************************************************************


'Self explaing Variable names
'Capabilities={3,4,7}
'CapabilityDescriptions={"Random Access","Supports Writing","Supports Removable Media"}
Public drvBytesPerSector As Integer '= 512
Public drvDescription   As String   '=Disk drive
Public drvDeviceID      As String   '=\\.\PHYSICALDRIVE1
Public drvIndex         As Integer  'indice nelle liste
Public drvMediaType     As String   '=Removable Media
Public drvModel         As String   '=SDXC Card
Public drvName          As String   '=\\.\PHYSICALDRIVE1
Public drvPartitions    As Integer  '= 2
Public drvSize          As Single     '196863851520#
Public drvStatus        As String   '= OK
Public drvTotalSectors  As Long     '= 384499710
Public drvStartSector   As Currency
Public drvNinFolder     As String
Public drvEmuFolder     As String

Public Partitions       As Collection   'of Partition type

Private abResult() As Byte
Private m_aSxEmuDataDisable(0 To 511)   As Byte
Private m_aSxEmuDataEnable(0 To 511)    As Byte
Private m_bSxEmuEnabled                 As Boolean
Private iDrvIndex                        As Integer

Const FILE_BEGIN = 0

Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)



''''Private Function MAKEWORD(ByVal bLo As Byte, ByVal bHi As Byte) As Integer
''''
''''    If bHi And &H80 Then
''''        MAKEWORD = (((bHi And &H7F) * 256) + bLo) Or &H8000
''''    Else
''''        MAKEWORD = (bHi * 256) + bLo
''''    End If
''''
''''End Function


Private Sub Class_Initialize()

    Dim idx         As Integer
    
    'Defaults
    drvBytesPerSector = 0
    drvDescription = ""
    drvDeviceID = ""
    drvIndex = -1
    drvMediaType = "NA"
    drvModel = "NA"
    drvName = "NA"
    drvPartitions = 0
    drvSize = 0
    drvStatus = "NA"
    drvTotalSectors = 0
    BytesPerSector = 512@
    iDrvIndex = 0
    
    'Initialize Array for Sx Emunand partition enabled/disabled
    For idx = 0 To 511
        m_aSxEmuDataDisable(idx) = 0
        m_aSxEmuDataEnable(idx) = 0
    Next

    'Set relevant bytes for Sx Partition Emu enabler
    m_aSxEmuDataEnable(0) = &H54
    m_aSxEmuDataEnable(1) = &H58
    m_aSxEmuDataEnable(2) = &H4E
    m_aSxEmuDataEnable(3) = &H41
    m_aSxEmuDataEnable(4) = &H4E
    m_aSxEmuDataEnable(5) = &H44
    'm_aSxEmuDataEnable(6,7,8,9,10, 11, 12, 13, 14, 15)=&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0,
    m_aSxEmuDataEnable(16) = &H54
    m_aSxEmuDataEnable(17) = &H65
    m_aSxEmuDataEnable(18) = &H61
    m_aSxEmuDataEnable(19) = &H6D
    m_aSxEmuDataEnable(20) = &H20
    m_aSxEmuDataEnable(21) = &H58
    m_aSxEmuDataEnable(22) = &H65
    m_aSxEmuDataEnable(23) = &H63
    m_aSxEmuDataEnable(24) = &H75
    m_aSxEmuDataEnable(25) = &H74
    m_aSxEmuDataEnable(26) = &H65
    m_aSxEmuDataEnable(27) = &H72
    m_aSxEmuDataEnable(28) = &H20
    m_aSxEmuDataEnable(29) = &H96
    m_aSxEmuDataEnable(30) = &H20
    m_aSxEmuDataEnable(31) = &H52
    m_aSxEmuDataEnable(32) = &H6F
    m_aSxEmuDataEnable(33) = &H63
    m_aSxEmuDataEnable(34) = &H6B
    m_aSxEmuDataEnable(35) = &H69
    m_aSxEmuDataEnable(36) = &H6E
    m_aSxEmuDataEnable(37) = &H67
    m_aSxEmuDataEnable(38) = &H20
    m_aSxEmuDataEnable(39) = &H74
    m_aSxEmuDataEnable(40) = &H68
    m_aSxEmuDataEnable(41) = &H65
    m_aSxEmuDataEnable(42) = &H20
    m_aSxEmuDataEnable(43) = &H53
    m_aSxEmuDataEnable(44) = &H77
    m_aSxEmuDataEnable(45) = &H69
    m_aSxEmuDataEnable(46) = &H74
    m_aSxEmuDataEnable(47) = &H63
    m_aSxEmuDataEnable(48) = &H68
    m_aSxEmuDataEnable(49) = &H20
    m_aSxEmuDataEnable(50) = &H66
    m_aSxEmuDataEnable(51) = &H6F
    m_aSxEmuDataEnable(52) = &H72
    m_aSxEmuDataEnable(53) = &H65
    m_aSxEmuDataEnable(54) = &H76
    m_aSxEmuDataEnable(55) = &H65
    m_aSxEmuDataEnable(56) = &H72
    m_aSxEmuDataEnable(57) = &H20
    m_aSxEmuDataEnable(58) = &H61
    m_aSxEmuDataEnable(59) = &H6E
    m_aSxEmuDataEnable(60) = &H64
    m_aSxEmuDataEnable(61) = &H20
    m_aSxEmuDataEnable(62) = &H62
    m_aSxEmuDataEnable(63) = &H65
    m_aSxEmuDataEnable(64) = &H79
    m_aSxEmuDataEnable(65) = &H6F
    m_aSxEmuDataEnable(66) = &H6E
    m_aSxEmuDataEnable(67) = &H64
    'm_aSxEmuDataEnable(68)=&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    m_aSxEmuDataEnable(81) = &H20
    m_aSxEmuDataEnable(82) = &HA4
    m_aSxEmuDataEnable(83) = &H3
    m_aSxEmuDataEnable(84) = &H2
    'm_aSxEmuDataEnable(85)=&H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0

End Sub

Public Property Get sxPartitionEnabled() As Boolean
    sxPartitionEnabled = m_bSxEmuEnabled
End Property

'Write the all 0 or the TX bytes to disable/enable the sxos emunand
Public Property Let sxPartitionEnabled(ByVal bSet As Boolean)

    Dim cbytes          As Long
    Dim res             As Long
    Dim aSdBuf()        As Byte
    Dim idx             As Integer

    cbytes = 512
    If bSet Then
        res = DirectWriteDriveNT(drvDeviceID, EMU_SX_ENABLE_SECTOR, m_aSxEmuDataEnable, cbytes)
        m_bSxEmuEnabled = True
        
    Else
        res = DirectWriteDriveNT(drvDeviceID, EMU_SX_ENABLE_SECTOR, m_aSxEmuDataDisable, cbytes)
        m_bSxEmuEnabled = False
        
    End If

    MsgBox LangLabel("IDidIt", "I did it!"), vbOKOnly, LangLabel("Done", "Done !")



End Property

'Set all properties based on the Disk object passed as arguments
Public Sub init(ByRef oInit As Variant)

    Dim col         As New Collection
    Dim Object
    Dim List
    Dim v
    
    Dim aSdBuf()    As Byte
    Dim res         As Long
    Dim idx         As Integer
    Dim cini        As cInifile

    If Not oInit Is Nothing Then
        drvBytesPerSector = oInit.BytesPerSector
        drvDescription = oInit.Description
        drvDeviceID = oInit.DeviceID
        drvIndex = -1
        iDrvIndex = oInit.Index
        drvMediaType = oInit.MediaType
        drvModel = oInit.Model
        drvName = oInit.Name
        drvPartitions = oInit.Partitions
        drvSize = oInit.Size
        drvStatus = oInit.Status
        drvTotalSectors = oInit.TotalSectors

'        : Name : "SectorsPerTrack" : String : clsDrive.init -----         : Name : "SectorsPerTrack" : String : clsDrive.init
'        : Name : "TracksPerCylinder" : String : clsDrive.init  --- Name : "TotalCylinders" : String : clsDrive.init    ---
'        : Name : "TotalTracks" : String : clsDrive.init -----         : Name : "TotalSectors" : String : clsDrive.init   -----         : Name : "TotalHeads" : String : clsDrive.init  ------



        'Find the drive letter
        'This piece of code is useful also for retrive partititon information about SD card
        Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_DiskDriveToDiskPartition")
        
        For Each Object In List
            
            If Replace(Replace(Split(Object.Antecedent, "=")(1), "\\", "\"), """", "") = drvDeviceID Then
                col.Add Replace(Split(Object.Dependent, "=")(1), """", ""), Replace(Split(Object.Dependent, "=")(1), """", "") '
            End If
        Next

        Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_LogicalDiskToPartition")
        For Each v In col
            For Each Object In List
                If Replace(Split(Object.Antecedent, "=")(1), """", "") = v Then drvName = Replace(Split(Object.Dependent, "=")(1), """", "")
            Next
        Next
    

        Set Partitions = New Collection

        Dim prt As clsPartition

        Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_DiskPartition")
        For Each Object In List
            If iDrvIndex = Object.DiskIndex Then
            Debug.Print "Object.Caption=" & Object.Caption
            Debug.Print "Object.Description=" & Object.Description
            Debug.Print "Object.DeviceID=" & Object.DeviceID
            Debug.Print "Object.DiskIndex=" & Object.DiskIndex
            Debug.Print "Object.Index=" & Object.Index
            Debug.Print "Object.Name=" & Object.Name
            Debug.Print "Object.NumberOfBlocks=" & Object.NumberOfBlocks
            Debug.Print "Object.Size=" & Object.Size
            Debug.Print "Object.StartingOffset=" & Object.StartingOffset
            Debug.Print "Object.Type=" & Object.Type

            Debug.Print " "

            Set prt = New clsPartition

            prt.DiskIndex = Object.DiskIndex
            prt.NumberOfBlocks = Object.NumberOfBlocks
            prt.PartitionIndex = Object.Index
            prt.StartingAddress = Object.StartingOffset
            prt.Size = Object.Size / 10000@
            prt.Name = Object.Name

            Partitions.Add prt, "P" & CStr(prt.PartitionIndex)

            'If Object.DeviceID = v Then drvName = Replace(Split(Object.Dependent, "=")(1), """", "")

            End If

        Next


    
    
    
    
    
    
    Else
        drvBytesPerSector = 0
        drvDescription = ""
        drvDeviceID = ""
        drvIndex = -1
        drvMediaType = "NA"
        drvModel = "NA"
        drvName = "NA"
        drvPartitions = 0
        drvSize = 0
        drvStatus = "NA"
        drvTotalSectors = 0
    
    End If
  
    'Check if SD is TX Emunand enabled
    res = DirectReadDriveNT(drvDeviceID, 0.0001@, aSdBuf, 512)
    
    If res > 0 Then
        'sufficienti i primi 85bytes
        For idx = 0 To 85
            If aSdBuf(idx) = m_aSxEmuDataEnable(idx) Then
                m_bSxEmuEnabled = True
            Else
                m_bSxEmuEnabled = False
                Exit For
            End If
        Next
    Else
        MsgBox "Unable to read from SD card, maybe I need Administrative rights to access your reader" & vbCrLf & _
                "Please restart the software using right click and Start as Administrator", vbOKOnly, "Error!"
    End If
    
    'If not an SXOS partition Emu check if there is an Atmo Emu by reading emummc.ini on SD
    If Not m_bSxEmuEnabled Then
    
        Set cini = New cInifile
        With cini
            .Path = drvName & "\emummc\emummc.ini"
            .Section = "emummc"
        
            .Key = "sector"
            .Default = -1
            
            'Questo "appoggia" il valore a destra, perchè la virgola è fittizia nei currency
            drvStartSector = CCur(Replace(.Value, "0x", "&h")) / 10000@
        
        
            .Key = "path"
            .Default = "\emummc"
            drvEmuFolder = Replace(.Value, "/", "\")
            
            .Key = "nintendo_path"
            .Default = "emuMMC/HPE0/Nintendo"
            drvNinFolder = Replace(.Value, "/", "\")
        
        End With
        Set cini = Nothing

    Else    'Se è una Emu sxos imposta i valori predefiniti.

        'imposto il settore a 2
        drvStartSector = 0.0002@
        drvEmuFolder = "sxos/Emunand"
        drvNinFolder = "Emutendo"
    
    End If


End Sub

'Read data from starting sector and write it to the file referred as lFileID (open) or HF (HugeBinaryFile)
Public Function readDataToFileId(ByRef lStartSect As Currency, ByVal lNumOfSect As Long, ByVal lFileID As Long, Optional ByRef HF As HugeBinaryFile = Nothing) As Long

    Dim iSecNum         As Integer
    Dim Buffer()        As Byte
    Dim res             As Long
    
    Dim SecToWrite      As Currency
    Dim RemainingSect   As Currency
    Dim WrittenSectors  As Currency
    Dim idx             As Long
    
    Dim MaxSector       As Currency
    MaxSector = 6.5535@ ' / 10000@
    
    
    On Error GoTo ERR_FUNCT
        
    RemainingSect = CCur(lNumOfSect) / 10000@
    WrittenSectors = 0
  
    
    Do
        If (RemainingSect > MaxSector) Then
            SecToWrite = MaxSector
            RemainingSect = RemainingSect - MaxSector
        Else
            SecToWrite = RemainingSect
            RemainingSect = 0
        End If

        res = DirectReadDriveNT(drvDeviceID, lStartSect + WrittenSectors, Buffer(), EMU_BLOCK_SIZE * (SecToWrite * 10000@))
        
        If res > 0 Then
        
            If HF Is Nothing Then
                Put lFileID, 1 + (WrittenSectors * EMU_BLOCK_SIZE * 10000), Buffer()
            Else
                HF.SeekAbsolute (WrittenSectors * EMU_BLOCK_SIZE * 10000)
                HF.WriteBytes Buffer()
            End If
            
            WrittenSectors = WrittenSectors + SecToWrite
        
            lblStatus.Caption = "Written Blocks: " & Format(CStr(WrittenSectors * 10000@), "#,###")
            DoEvents
        
            If g_bStopOperations Then Exit Function

        Else
            MsgBox LangLabel("ErroreLettura", "Error reading disk, start EmuTool in Administrative mode")
            Exit Do
        End If
        
    Loop While RemainingSect > 0
    
    
       
EXIT_FUNCT:
    Exit Function
       
ERR_FUNCT:

    If Err.Number = 1 Then
        Resume
    Else
        MsgBox LangLabel("Errore", "Errore") & ": " & Err.Number & " - " & Err.Description
        Resume EXIT_FUNCT
    End If
       
End Function

'Procedure to read data from a HF file and copy to the drive startinf from the specified sector
Public Function readDataFromFileId(ByRef lStartSect As Currency, ByRef HF As HugeBinaryFile) As Long


    Dim iSecNum         As Integer
    Dim Buffer()        As Byte
    Dim res             As Long

    Dim SecToWrite      As Currency
    Dim RemainingSect   As Currency
    Dim WrittenSectors  As Currency
    Dim idx             As Long

    Dim MaxSector       As Currency
    MaxSector = 6.5535@ ' / 10000@


    On Error GoTo ERR_FUNCT

    RemainingSect = (HF.FileLen / EMU_BLOCK_SIZE) / 10000@
    WrittenSectors = 0


    Do
        If (RemainingSect > MaxSector) Then
            SecToWrite = MaxSector
            RemainingSect = RemainingSect - MaxSector
        Else
            SecToWrite = RemainingSect
            RemainingSect = 0
        End If

        'res = DirectReadDriveNT(drvDeviceID, lStartSect + WrittenSectors, Buffer(), EMU_BLOCK_SIZE * (SecToWrite * 10000@))
        '
        ReDim Buffer(EMU_BLOCK_SIZE * (SecToWrite * 10000@) - 1)
        HF.SeekAbsolute (WrittenSectors * EMU_BLOCK_SIZE * 10000@)
        res = HF.ReadBytes(Buffer)

        If res > 0 Then
            
            res = DirectWriteDriveNT(drvDeviceID, lStartSect + WrittenSectors, Buffer(), SecToWrite * EMU_BLOCK_SIZE * 10000@)

            WrittenSectors = WrittenSectors + SecToWrite

            lblStatus.Caption = "Written Blocks: " & Format(CStr(WrittenSectors * 10000@), "#,###")
            DoEvents

            If g_bStopOperations Then Exit Function

        Else
            MsgBox LangLabel("ErroreLettura", "Error reading disk, start EmuTool in Administrative mode")
            Exit Do
        End If

    Loop While RemainingSect > 0


       
EXIT_FUNCT:
    Exit Function
       
ERR_FUNCT:

    If Err.Number = 1 Then
        Resume
    Else
        MsgBox LangLabel("Errore", "Errore") & ": " & Err.Number & " - " & Err.Description
        Resume EXIT_FUNCT
    End If

End Function


