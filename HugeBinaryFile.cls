VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HugeBinaryFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'HugeBinaryFile
'==============
'Based on an old MS KB article for VB4
'
'A class for doing simple binary I/O on very large disk files
'(well over the usual 2GB limit).  It only does I/O using Byte
'arrays, and makes use of Currency values that are scaled to
'whole numbers in places:
'
'   For a file of one byte the FileLen property returns 1.0000 as
'   its value.
'
'Operation is similar in many ways to native VB Get#/Put# I/O, for
'example the EOF property must be checked after a ReadBytes() call.
'You must also Dim/Redim buffers to desired sizes before calling
'ReadBytes() or WriteBytes().
'
'Short (signed Long) relative seeks and long (unsigned Currency)
'absolute seeks from 0 may be done.
'
'AutoFlush may be set True to force buffer flushes on every write.
'The Flush() method may be called explicitly if necessary.
'
'*************************************************
'*
'*  Modified by TheyKilledKenny on 17 Oct 2019
'*
'*************************************************



Public Enum HBF_Errors
  HBF_UNKNOWN_ERROR = 45600
  HBF_FILE_ALREADY_OPEN
  HBF_OPEN_FAILURE
  HBF_SEEK_FAILURE
  HBF_FILELEN_FAILURE
  HBF_READ_FAILURE
  HBF_WRITE_FAILURE
  HBF_FILE_ALREADY_CLOSED
End Enum

Private Const HBF_SOURCE = "EmuTool"

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    lpSource As Long, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Any) As Long

Private hFile As Long
Private sFName As String
Private fAutoFlush As Boolean
Private fEOF As Boolean
Private C As TKK_Cur
Private L As Cur2Long

Public Property Get AutoFlush() As Boolean
    RaiseErrorIfClosed
    AutoFlush = fAutoFlush
End Property

Public Property Let AutoFlush(ByVal NewVal As Boolean)
    RaiseErrorIfClosed
    fAutoFlush = NewVal
End Property

Public Property Get FileHandle() As Long
    RaiseErrorIfClosed
    FileHandle = hFile
End Property

Public Property Get FileLen() As Currency
    RaiseErrorIfClosed
    L.LowVal = GetFileSize(hFile, L.HighVal)
    If L.LowVal = INVALID_FILE_SIZE Then
        If Err.LastDllError Then RaiseError HBF_FILELEN_FAILURE
        L.LowVal = 0
        L.HighVal = 0
    End If
    LSet C = L
    FileLen = C.Value * 10000@
End Property

Public Property Get FileName() As String
    RaiseErrorIfClosed
    FileName = sFName
End Property

Public Property Get EOF() As Boolean
    RaiseErrorIfClosed
    EOF = fEOF
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = hFile <> INVALID_HANDLE_VALUE
End Property

Public Sub CloseFile()
    RaiseErrorIfClosed
    CloseHandle hFile
    sFName = ""
    fAutoFlush = False
    fEOF = False
    hFile = INVALID_HANDLE_VALUE
End Sub

Public Sub Flush()
    RaiseErrorIfClosed
    FlushFileBuffers hFile
End Sub

'Added the flag to open a file only if it Exist (not create)
Public Function OpenFile(ByVal OpenFileName As String, Optional ByVal bOpenExisting As Boolean = False) As Boolean
    
    If hFile <> INVALID_HANDLE_VALUE Then
        MsgBox RaiseError(HBF_FILE_ALREADY_OPEN), vbOKOnly, LangLabel("Errore", "Errore!")
        OpenFile = False
    Else
    
        If bOpenExisting Then
            hFile = CreateFile(OpenFileName, GENERIC_WRITE Or GENERIC_READ, 0, _
                               0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        Else
            hFile = CreateFile(OpenFileName, GENERIC_WRITE Or GENERIC_READ, 0, _
                               0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        End If
        
        
        If hFile = INVALID_HANDLE_VALUE Then
            sFName = RaiseError(HBF_OPEN_FAILURE) ', vbOKOnly, LangLabel("Errore", "Errore!")
            OpenFile = False
        Else
            OpenFile = True
            sFName = OpenFileName
        End If
    
    End If
    
    
EXIT_SUB:

    Exit Function

ERR_SUB:

    MsgBox Err.Number & " - " & Err.Description & "(" & OpenFileName & ")"
    Resume EXIT_SUB
    
End Function

Public Function ReadBytes(ByRef Buffer() As Byte) As Long
    RaiseErrorIfClosed
    If ReadFile(hFile, _
                Buffer(LBound(Buffer)), _
                UBound(Buffer) - LBound(Buffer) + 1, _
                ReadBytes, _
                0) Then
        If ReadBytes = 0 Then
            fEOF = True
        End If
    Else
        RaiseError HBF_READ_FAILURE
    End If
End Function

Public Sub SeekAbsolute(ByVal Position As Currency)
    RaiseErrorIfClosed
    C.Value = Position / 10000@
    LSet L = C
    If SetFilePointer(hFile, L.LowVal, L.HighVal, FILE_BEGIN) _
        = INVALID_SET_FILE_POINTER Then
            If Err.LastDllError Then RaiseError HBF_SEEK_FAILURE
    End If
End Sub

Public Sub SeekEnd()
    RaiseErrorIfClosed
    If SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) _
        = INVALID_SET_FILE_POINTER Then
            RaiseError HBF_SEEK_FAILURE
    End If
End Sub

Public Sub SeekRelative(ByVal Offset As Long)
    'Offset is signed.
    RaiseErrorIfClosed
    If SetFilePointer(hFile, Offset, ByVal 0&, FILE_CURRENT) _
        = INVALID_SET_FILE_POINTER Then
            RaiseError HBF_SEEK_FAILURE
    End If
End Sub

Public Function WriteBytes(Buffer() As Byte) As Long
    RaiseErrorIfClosed
    If WriteFile(hFile, _
                 Buffer(LBound(Buffer)), _
                 UBound(Buffer) - LBound(Buffer) + 1, _
                 WriteBytes, _
                 0) Then
        If fAutoFlush Then Flush
    Else
        RaiseError HBF_WRITE_FAILURE
    End If
End Function

Private Sub Class_Initialize()
    hFile = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
    If hFile <> INVALID_HANDLE_VALUE Then CloseHandle hFile
End Sub

Private Function RaiseError(ByVal ErrorCode As HBF_Errors) As String

    Dim Win32Err As Long, Win32Text As String

    Win32Err = Err.LastDllError
    If Win32Err Then
        Win32Text = vbNewLine & "Error " & Win32Err & vbNewLine _
                  & DecodeAPIErrors(Win32Err)
    End If
    If IsOpen Then CloseFile
    
    On Error GoTo ERR_SUB
    
    Select Case ErrorCode
        Case HBF_FILE_ALREADY_OPEN
            RaiseError = CStr(HBF_FILE_ALREADY_OPEN) & "- File already open "
        Case HBF_OPEN_FAILURE
            RaiseError = CStr(HBF_OPEN_FAILURE) & " - Error opening file. " & Win32Text
        Case HBF_SEEK_FAILURE
            RaiseError = CStr(HBF_SEEK_FAILURE) & " - Seek Error. " & Win32Text
        Case HBF_FILELEN_FAILURE
            RaiseError = CStr(HBF_FILELEN_FAILURE) & " - GetFileSize Error. " & Win32Text
        Case HBF_READ_FAILURE
            RaiseError = CStr(HBF_READ_FAILURE) & " - Read failure. " & Win32Text
        Case HBF_WRITE_FAILURE
            RaiseError = CStr(HBF_WRITE_FAILURE) & " - Write failure. " & Win32Text
        Case HBF_FILE_ALREADY_CLOSED
            RaiseError = CStr(HBF_FILE_ALREADY_CLOSED) & " - File must be open for this operation"
        Case Else
            RaiseError = CStr(HBF_UNKNOWN_ERROR) & " - Unknown error " & Win32Text
    End Select
    
    
EXIT_SUB:

    Exit Function
    
ERR_SUB:

    
End Function

Private Sub RaiseErrorIfClosed()
    If hFile = INVALID_HANDLE_VALUE Then RaiseError HBF_FILE_ALREADY_CLOSED
End Sub

Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
    Dim strMsg As String, lngMsgLen As Long

    strMsg = Space$(256)
    lngMsgLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                              ErrorCode, 0&, strMsg, 256&, 0&)
    If lngMsgLen > 0 Then
        DecodeAPIErrors = Left(strMsg, lngMsgLen)
    Else
        DecodeAPIErrors = "Unknown Error."
    End If
End Function




'*************************************************
'*
'*  Added by TheyKilledKenny on 17 Oct 2019
'*
'*************************************************
Public Function readDataToFileId(ByRef lStartByte As Currency, ByVal lNumOfBytes As Currency, ByVal cDstStartByte As Currency, ByRef HF As HugeBinaryFile) As Currency

    Dim iSecNum         As Integer
    Dim Buffer()        As Byte
    Dim res             As Long
    
    Dim BytesToWrite      As Currency   'SecToWrite
    Dim RemainingBytes  As Currency
    Dim WrittenSectors  As Currency
    Dim idx             As Long
    
    Dim MaxSector       As Currency
    MaxSector = 65535@ * 512@   ' / 10000@
    
    
    On Error GoTo ERR_FUNCT
        
    RemainingBytes = lNumOfBytes
    WrittenSectors = 0
  
    
    Do
        If (RemainingBytes > MaxSector) Then
            BytesToWrite = MaxSector
            RemainingBytes = RemainingBytes - MaxSector
        Else
            BytesToWrite = RemainingBytes
            RemainingBytes = 0
        End If

        'res = DirectReadDriveNT(drvDeviceID, lStartSect + WrittenSectors, Buffer(), EMU_BLOCK_SIZE * (SecToWrite * 10000@))
        ReDim Buffer(BytesToWrite - 1)
        
        SeekAbsolute (lStartByte + WrittenSectors)
        res = ReadBytes(Buffer())
        
        
        If res > 0 Then
        
            HF.SeekAbsolute (cDstStartByte + WrittenSectors)
            HF.WriteBytes Buffer()
            
            WrittenSectors = WrittenSectors + BytesToWrite
        
            lblStatus.Caption = "Written blocks: " & Format(CStr(WrittenSectors / 512), "#,###")
            DoEvents
        
            If g_bStopOperations Then Exit Function

        Else
            MsgBox LangLabel("ErroreLettura", "Error reading disk, start EmuTool in Administrative mode")
            Exit Do
        End If
        
    Loop While RemainingBytes > 0
    
    
       
EXIT_FUNCT:

    readDataToFileId = WrittenSectors
    Exit Function
       
ERR_FUNCT:

'    If Err.Number = 1 Then
'        Resume
'    Else
        MsgBox LangLabel("Errore", "Errore") & ": " & Err.Number & " - " & Err.Description
        Resume EXIT_FUNCT
'    End If
       
End Function
