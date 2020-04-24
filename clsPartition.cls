VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPartition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public DiskIndex As Integer
Public PartitionIndex As Integer
Public StartingAddress As Currency
Public NumberOfBlocks As Currency
Public Size As Currency
Public Name As String
            
        
Private Sub Class_Initialize()

    DiskIndex = -1
    PartitionIndex = -1
    StartingAddress = -1
    NumberOfBlocks = -1

End Sub
