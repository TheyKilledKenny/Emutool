Attribute VB_Name = "diskio"
'*****************************************************
'*
'* modified by TheyKilledKenny on 28 Sept 2019
'*
'* rewritten the main functions DirectReadDriveNT and DirectWriteDriveNT
'* because Long datatype is limited. Converted to Currency as it is a plain 64bit with virtual fixed comma
'*
'* Drive information and Read/write procedures
'*
'*****************************************************

'Rem 8-11-2008 'ByRef lpBuffer' and 'cBytes' long return modification
'Rem by Erdogan Tan
'Rem ***
'Rem "DirectReadDriveNT" function was originally written
'Rem by Arkadiy Olovyannikov
'Rem with a variant return
'Rem ... for reading logical (dos/windows) drive/disk sectors...
'Rem ***
'Rem Physical disk read/write features/procedures is written by Erdogan Tan
'Rem by using information on Microsoft Developers Network (MSDN) web site
'Rem http://msdn.microsoft.com/tr-tr/library/aa363858(en-us,VS.85).aspx
'Rem Adapted to Visual Basic (5.0) code by Erdogan Tan on 27-10-2008
'Rem ... and successfully realized on Windows XP SP3 (8-11-2008)
'Rem This code is successfully running on Windows XP Home & Professional.

'**********************************************************
'VB Project FILE: diskio.bas  (diskio module)
'************ VB 5.0/6.0 Source Code **********************

'*****************************************************************
' Module for performing Direct Read/Write access to disk sectors
'
' Written by Arkadiy Olovyannikov (ark@fesma.ru)
'*****************************************************************

'*************Win9x direct Read/Write Staff**********

Public Enum FAT_WRITE_AREA_CODE
    FAT_AREA = &H2001
    ROOT_DIR_AREA = &H4001
    DATA_AREA = &H6001
End Enum

Public Type DISK_IO
  dwStartSector As Long
  wSectors As Integer
  dwBuffer As Long
End Type


Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const INVALID_HANDLE_VALUE = -1&

Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Public Const CREATE_ALWAYS = 2
Public Const OPEN_ALWAYS = 4
Public Const INVALID_SET_FILE_POINTER = -1
Public Const INVALID_FILE_SIZE = -1

Public Const FILE_BEGIN = 0, FILE_CURRENT = 1, FILE_END = 2


'''''Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
''Public Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, _
''                                                        ByRef lpInBuffer As Any, _
''                                                        ByVal nInBufferSize As Long, _
''                                                        ByRef lpOutBuffer As Any, _
''                                                        ByVal nOutBufferSize As Long, _
''                                                        ByRef lpBytesReturned As Long, _
''                                                        ByVal lpOverlapped As Long) As Long
                                                        
''Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As SafeFileHandle, ByVal dwIoControlCode As Long, _
''                                                        ByRef lpInBuffer As SENDCMDINPARAMS, _
''                                                        ByVal nInBufferSize As Long, _
''                                                        ByRef lpOutBuffer As SENDCMDOUTPARAMS, _
''                                                        ByVal nOutBufferSize As Long, _
''                                                        ByRef lpBytesReturned As Long, _
''                                                        ByVal lpOverlapped As Long) As Long

Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Public Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'DRIVE_UNKNOWN = 0 '    The drive type cannot be determined.
'DRIVE_NO_ROOT_DIR = 1 ' The root path is invalid; for example, there is no volume mounted at the specified path.
'DRIVE_REMOVABLE = 2 ' The drive has removable media; for example, a floppy drive, thumb drive, or flash card reader.
'DRIVE_FIXED = 3     ' The drive has fixed media; for example, a hard disk drive or flash drive.
'DRIVE_REMOTE = 4    ' The drive is a remote (network) drive.
'DRIVE_CDROM = 5     ' The drive is a CD-ROM drive.
'DRIVE_RAMDISK = 6   ' The drive is a RAM disk.


Public Type TKK_Cur
    Value As Currency
End Type

Public Type Cur2Long
    LowVal As Long
    HighVal As Long
End Type

Private C As TKK_Cur
Private L As Cur2Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Additions by Erdogan Tan
Public BytesPerSector As Currency

'Removed By TheyKilledKenny, not needed LoWord and HiWord doesn't work with Currency or double number (Mod does not work)

'''''''Private Function IsWindowsNT() As Boolean
'''''''   Dim verinfo As OSVERSIONINFO
'''''''   verinfo.dwOSVersionInfoSize = Len(verinfo)
'''''''   If (GetVersionEx(verinfo)) = 0 Then Exit Function
'''''''   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
'''''''End Function

''''''''Private Function LoWord(ByVal dNum As Double) As Long
''''''''    LoWord = CLng(dNum Mod 65536)
''''''''End Function

''''''''Private Function HiWord(ByVal dNum As Double) As Long
''''''''    HiWord = CLng(dNum - (dNum Mod 65536))
''''''''End Function
''''''''

''=============NT staff=============
''Read/Write drive with any file system
'
'Rem 8-11-2008 'ByRef lpBuffer' and 'cBytes' long return modification
'Rem by Erdogan Tan
'Rem ***
'Rem "DirectReadDriveNT" function was originally written
'Rem by Arkadiy Olovyannikov
'Rem with a variant return
'Rem ... for reading logical (dos/windows) drive/disk sectors...
'Rem ***
'Rem Physical disk read/write features/procedures is written by Erdogan Tan
'Rem by using information on Microsoft Developers Network (MSDN) web site
'Rem http://msdn.microsoft.com/tr-tr/library/aa363858(en-us,VS.85).aspx
'Rem Adapted to Visual Basic (5.0) code by Erdogan Tan on 27-10-2008
'Rem ... and successfully realized on Windows XP SP3 (8-11-2008)
'Rem This code is successfully running on Windows XP Home & Professional.
'''''''
''''''''''''Public Function DirectReadDriveNT(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByRef lpBuffer() As Byte, ByVal cbytes As Long) As Long
''''''''''''
''''''''''''    Dim hDevice As Long
''''''''''''    Dim abBuff() As Byte
''''''''''''    Dim nSectors As Integer
''''''''''''
''''''''''''    nSectors = Int((iOffset + cbytes - 1) / BytesPerSector) + 1
''''''''''''
''''''''''''    Rem hDevice = CreateFile("\\.\" & UCase(Left(sDrive, 1)) & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
''''''''''''    Rem 4-11-2008 Physical disk read/write modification
''''''''''''
''''''''''''    hDevice = CreateFile(sDrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
''''''''''''    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
''''''''''''    Call SetFilePointer(hDevice, iStartSec * BytesPerSector, 0, FILE_BEGIN)
''''''''''''    ReDim lpBuffer(cbytes - 1)
''''''''''''    ReDim abBuff(nSectors * BytesPerSector - 1)
''''''''''''    Call ReadFile(hDevice, abBuff(0), UBound(abBuff) + 1, cbytes, 0&)
''''''''''''    CloseHandle hDevice
''''''''''''    CopyMemory lpBuffer(0), abBuff(iOffset), cbytes
''''''''''''    DirectReadDriveNT = cbytes
''''''''''''
''''''''''''End Function

'Rewritten by TheyKiledKenny, the original one doesn't work for overflow on high ector number
Public Function DirectReadDriveNT(ByVal sDrive As String, ByVal iStartSec As Currency, ByRef lpBuffer() As Byte, ByVal cbytes As Long) As Long
  
    Dim hDevice     As Long
    Dim abBuff()    As Byte
    Dim nSectors    As Integer
    
    
    Debug.Print GetDriveType("c:")
    Debug.Print GetDriveType("f:")
    Debug.Print GetDriveType("g:")
    Debug.Print GetDriveType("d:")
    Debug.Print GetDriveType("t:")
    
    
    
    hDevice = CreateFile(sDrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    
    If hDevice = INVALID_HANDLE_VALUE Then Exit Function
    
    C.Value = iStartSec * BytesPerSector
    LSet L = C
    Call SetFilePointer(hDevice, L.LowVal, L.HighVal, FILE_BEGIN)
    
    ReDim lpBuffer(cbytes - 1)
    
    Call ReadFile(hDevice, lpBuffer(0), UBound(lpBuffer) + 1, cbytes, 0&)
    
    CloseHandle hDevice
    
    DirectReadDriveNT = cbytes
    
End Function

Public Function DirectWriteDriveNT(ByVal sDrive As String, ByVal iStartSec As Currency, ByRef lpBuffer() As Byte, ByVal cbytes As Long) As Boolean

    Dim hDevice     As Long
    Dim abBuff()    As Byte
    Dim ab()        As Byte
    Dim nRead       As Long
    
    hDevice = CreateFile(sDrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)

    If hDevice = INVALID_HANDLE_VALUE Then
        DirectWriteDriveNT = False
        Exit Function
    End If

    C.Value = iStartSec * BytesPerSector
    LSet L = C
    Call SetFilePointer(hDevice, L.LowVal, L.HighVal, FILE_BEGIN)

    Call LockFile(hDevice, L.LowVal, L.HighVal, cbytes, 0)

    DirectWriteDriveNT = WriteFile(hDevice, lpBuffer(0), UBound(lpBuffer) + 1, nRead, 0&)

    Call FlushFileBuffers(hDevice)
    Call UnlockFile(hDevice, L.LowVal, L.HighVal, cbytes, 0)

    CloseHandle hDevice

End Function

'Deleted and Rewritten by TheyKilledKenny
'
'''''''Public Function DirectWriteDriveNT(ByVal sDrive As String, ByVal iStartSec As Long, ByVal iOffset As Long, ByRef lpBuffer() As Byte, ByVal cBytes As Long) As Long
'''''''
'''''''    Rem Dim hDevice As Long
'''''''    Dim nSectors As Long
'''''''    Dim nRead As Long
'''''''    Rem Dim abBuff() As Byte
'''''''
'''''''    nSectors = Int((iOffset + cBytes - 1) / BytesPerSector) + 1
'''''''    Rem hDevice = CreateFile("\\.\" & UCase(Left(sDrive, 1)) & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
'''''''    Rem 4-11-2008 Physical disk read/write modification
'''''''    Rem hDevice = CreateFile("\\.\" & sDrive, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
'''''''    Rem If hDevice = INVALID_HANDLE_VALUE Then Exit Function
'''''''    ReDim abBuff(nSectors * BytesPerSector - 1) As Byte
'''''''    Call DirectReadDriveNT(sDrive, iStartSec, 0, abBuff(), nSectors * BytesPerSector)
'''''''    CopyMemory abBuff(iOffset), lpBuffer(0), cBytes
'''''''    Rem 11-11-2008 Fixing 2 GB (32 bit - 64 bit) problem (to pass beyond 2 GB)
'''''''    Rem If Fix64BitNumber(iStartSec, CLng(BytesPerSector), lDistanceToMove, lpDistanceToMoveHigh) <> True Then Exit Function
'''''''    Call Fix64BitNumber(iStartSec, BytesPerSector, lDistanceToMove, lpDistanceToMoveHigh)
'''''''    If lDistanceToMove <= &H7FFFFFFF Then
'''''''       Call SetFilePointer(hDevice, CLng(lDistanceToMove), CLng(lpDistanceToMoveHigh), FILE_BEGIN)
'''''''    Else
'''''''       Call SetFilePointer(hDevice, &H7FFFFFFF, CLng(lpDistanceToMoveHigh), FILE_BEGIN)
'''''''       lDistanceToMove = lDistanceToMove - &H7FFFFFFF
'''''''       Call SetFilePointer(hDevice, CLng(lDistanceToMove), 0, FILE_CURRENT)
'''''''    End If
'''''''    Call LockFile(hDevice, CLng(lDistanceToMove), CLng(lpDistanceToMoveHigh), LoWord(nSectors * BytesPerSector), HiWord(nSectors * BytesPerSector))
'''''''    Call WriteFile(hDevice, abBuff(0), UBound(abBuff) + 1, nRead, 0&)
'''''''    Call FlushFileBuffers(hDevice)
'''''''    Call UnlockFile(hDevice, CLng(lDistanceToMove), CLng(lpDistanceToMoveHigh), LoWord(nSectors * BytesPerSector), HiWord(nSectors * BytesPerSector))
'''''''    Rem CloseHandle hDevice
'''''''    Rem 8-11-2008 Count of written bytes is equal to the required byte count?
'''''''    If nRead = UBound(abBuff) + 1 Then
'''''''       Rem All of the requested bytes are written...
'''''''       DirectWriteDriveNT = cBytes
'''''''    Rem Else
'''''''       Rem DirectWriteDriveNT = 0
'''''''    End If
'''''''End Function
