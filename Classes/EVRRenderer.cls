VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EVRRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' IMFVideoDisplayControl
Const IID_IMFVideoDisplayCtrl = "{A490B1E4-AB84-4D31-A1B2-181E03B1077A}"

Const IID_IMFGetServices = "{FA993888-4383-415A-A930-DD472A8CF6F7}"

Private Type GUID

    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(7) As Byte

End Type

Private byteGUID(14) As Byte

Private Declare Function CLSIDFromString _
                Lib "ole32" (ByVal lpszCLSID As Long, _
                             clsid As GUID) As Long

Private ifPlayback As olelib.IUnknown

Private ifServices As olelib.IUnknown

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Public Sub CreateInterface(baseFilter As olelib.IUnknown)

    Dim uuidGetServices As UUID, uuidVideoDisplayCtrl As GUID

    Dim result          As Long, b As Variant
    
    olelib.CLSIDFromString IID_IMFGetServices, uuidGetServices
    CLSIDFromString byteGUID(0), uuidVideoDisplayCtrl
    
    result = baseFilter.QueryInterface(uuidGetServices, ifServices)
    
    If (result <> 0) Then Err.Raise result
     
End Sub

