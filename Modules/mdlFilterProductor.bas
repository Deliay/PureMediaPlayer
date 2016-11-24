Attribute VB_Name = "mdlFilterProductor"
Option Explicit


Private Declare Function LAVAudioInstance _
                Lib "LAVAudio.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, b As olelib.UUID, ByRef c As olelib.IUnknown) As Long

Private Declare Function LAVSplitterInstance _
                Lib "LAVSplitter.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, b As olelib.UUID, ByRef c As olelib.IUnknown) As Long

Private Declare Function LAVVideoInstance _
                Lib "LAVVideo.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, b As olelib.UUID, ByRef c As olelib.IUnknown) As Long

Private Declare Function VSFilterInstance _
                Lib "vsfilter.dll" _
                Alias "DllGetClassObject" (a As olelib.UUID, b As olelib.UUID, ByRef c As olelib.IUnknown) As Long

Private Declare Function MadVRInstance _
                Lib "madVR.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, b As olelib.UUID, ByRef c As olelib.IUnknown) As Long

Private uuidIClassFactory As olelib.UUID
Private uuidIClassFactoryInited As Boolean
Private uuidIUnknown As olelib.UUID
Private uuidIUnknownInited As Boolean
Private objGlobalFilManager As FilgraphManager

Private Function getClassFactoryUUID() As olelib.UUID
    If (uuidIClassFactoryInited = False) Then
        '00000001-0000-0000-C000-000000000046
        DEFINE_GUID uuidIClassFactory, &H1, &H0, &H0, &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46
    End If
    getClassFactoryUUID = uuidIClassFactory
    
End Function

Private Function getUnknownUUID() As olelib.UUID
    If (uuidIUnknownInited = False) Then
        '00000000-0000-0000-C000-000000000046
        DEFINE_GUID uuidIUnknown, &H0, &H0, &H0, &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46
    End If
    getUnknownUUID = uuidIUnknown
End Function

Public Function DEFINE_GUID(ByRef ruuid As olelib.UUID, ByVal a As Long, ByVal b As Integer, ByVal c As Integer, ByVal d As Byte, ByVal e As Byte, ByVal f As Byte, ByVal g As Byte, ByVal h As Byte, ByVal i As Byte, ByVal j As Byte, ByVal k As Byte) As olelib.UUID
    With ruuid
        .Data1 = a
        .Data2 = b
        .Data3 = c
        .Data4(0) = d
        .Data4(1) = e
        .Data4(2) = f
        .Data4(3) = g
        .Data4(4) = h
        .Data4(5) = i
        .Data4(6) = j
        .Data4(7) = k
        
    End With
End Function

Public Function GetLAVAudioInstance() As IBaseFilter

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory
    '{E8E73B6B-4CB3-44A4-BE99-4F7BCB96E491}
    DEFINE_GUID uuidSrc, &HE8E73B6B, &H4CB3, &H44A4, &HBE, &H99, &H4F, &H7B, &HCB, &H96, &HE4, &H91
    
    If (LAVAudioInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetLAVAudioInstance
    End If
    
End Function

Public Function GetLAVVideoInstance() As IBaseFilter

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory
    'EE30215D-164F-4A92-A4EB-9D4C13390F9F
    DEFINE_GUID uuidSrc, &HEE30215D, &H164F, &H4A92, &HA4, &HEB, &H9D, &H4C, &H13, &H39, &HF, &H9F
    
    If (LAVVideoInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetLAVVideoInstance
    End If
    
End Function

Public Function GetLAVSplitterInstance() As IBaseFilter

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory
    '171252A0-8820-4AFE-9DF8-5C92B2D66B04
    DEFINE_GUID uuidSrc, &H171252A0, &H8820, &HA4FE, &H9D, &HF8, &H5C, &H92, &HB2, &HD6, &H6B, &H4
    
    If (LAVSplitterInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetLAVSplitterInstance
    End If
    
End Function


Public Function GetVSFilterInstance() As IBaseFilter

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory
    '93A22E7A-5091-45ef-BA61-6DA26156A5D0
    DEFINE_GUID uuidSrc, &H93A22E7A, &H5091, &H45EF, &HBA, &H61, &H6D, &HA2, &H61, &H56, &HA5, &HD0
    
    If (VSFilterInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetVSFilterInstance
    End If
    
End Function

Public Function GetMadVRInstance() As IBaseFilter

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory
    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    DEFINE_GUID uuidSrc, &HE1A8B82A, &H32CE, &H4B0D, &H4B, &HD, &HAA, &H68, &HC7, &H72, &HE4, &H23
    
    If (MadVRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetMadVRInstance
    End If
    
End Function

Private Function GetIGraphBuilder(Optional objFilManager As FilgraphManager = Nothing) As IGraphBuilder
    If (objFilManager = Nothing) Then
        Set GetIGraphBuilder = objGlobalFilManager
    Else
        Set GetIGraphBuilder = objFilManager
    End If
End Function

Public Function BuildGraph(ByVal strMediaFile) As FilgraphManager
    If (Not objGlobalFilManager Is Nothing) Then
        objGlobalFilManager.Pause
        objGlobalFilManager.Stop
        
    End If
    
    Set objGlobalFilManager = New FilgraphManager
    
    Dim objGraph As IGraphBuilder
    Dim objSplitter As IBaseFilter, objAudio As IBaseFilter, objVideo As IBaseFilter, objSubtitle As IBaseFilter, objRender As IBaseFilter
    
    
    Set objGraph = objGlobalFilManager
    Set objSplitter = GetLAVSplitterInstance
    objGraph.AddFilter objSplitter, "Splitter"
    
    
End Function
