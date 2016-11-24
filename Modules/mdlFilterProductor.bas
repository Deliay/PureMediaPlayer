Attribute VB_Name = "mdlFilterProductor"
Option Explicit

Private Declare Function LAVAudioInstance _
                Lib "LAVAudio.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, _
                                           b As olelib.UUID, _
                                           ByRef c As olelib.IUnknown) As Long

Private Declare Function LAVSplitterInstance _
                Lib "LAVSplitter.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, _
                                           b As olelib.UUID, _
                                           ByRef c As olelib.IUnknown) As Long

Private Declare Function LAVVideoInstance _
                Lib "LAVVideo.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, _
                                           b As olelib.UUID, _
                                           ByRef c As olelib.IUnknown) As Long

Private Declare Function VSFilterInstance _
                Lib "vsfilter.dll" _
                Alias "DllGetClassObject" (a As olelib.UUID, _
                                           b As olelib.UUID, _
                                           ByRef c As olelib.IUnknown) As Long

Private Declare Function MadVRInstance _
                Lib "madVR.ax" _
                Alias "DllGetClassObject" (a As olelib.UUID, _
                                           b As olelib.UUID, _
                                           ByRef c As olelib.IUnknown) As Long

Private Declare Function VRInstance _
                Lib "Quartz.dll" _
                Alias "DllGetClassObject" (a As olelib.UUID, _
                                           b As olelib.UUID, _
                                           ByRef c As olelib.IUnknown) As Long

Private uuidIClassFactory       As olelib.UUID

Private uuidIClassFactoryInited As Boolean

Private uuidIUnknown            As olelib.UUID

Private uuidIUnknownInited      As Boolean

Private objGlobalFilManager     As FilgraphManager

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

Public Function DEFINE_GUID(ByRef ruuid As olelib.UUID, _
                            ByVal a As Long, _
                            ByVal b As Integer, _
                            ByVal c As Integer, _
                            ByVal d As Byte, _
                            ByVal e As Byte, _
                            ByVal f As Byte, _
                            ByVal g As Byte, _
                            ByVal h As Byte, _
                            ByVal i As Byte, _
                            ByVal j As Byte, _
                            ByVal k As Byte) As olelib.UUID

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

Public Function GetVMR9Instance() As IBaseFilter

    ' &H51b4abf3, &H748f, &H4e3b, &Ha2, &H76, &Hc8, &H28, &H33, &H0e, &H92, &H6a
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    DEFINE_GUID uuidSrc, &H51B4ABF3, &H748F, &H4E3B, &HA2, &H76, &HC8, &H28, &H33, &HE, &H92, &H6A
    
    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetVMR9Instance
    End If
    
End Function

Public Function GetVMR7Instance() As IBaseFilter

    '&Hb87beb7b, &H8d29, &H423f, &Hae, &H4d, &H65, &H82, &Hc1, &H01, &H75, &Hac
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    DEFINE_GUID uuidSrc, &HB87BEB7B, &H8D29, &H423F, &HAE, &H4D, &H65, &H82, &HC1, &H1, &H75, &HAC
    
    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetVMR7Instance
    End If

End Function

Public Function GetVRInstance() As IBaseFilter

    '&H6bc1cffa, &H8fc1, &H4261, &Hac, &H22, &Hcf, &Hb4, &Hcc, &H38, &Hdb, &H50
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    DEFINE_GUID uuidSrc, &H6BC1CFFA, &H8FC1, &H4261, &HAC, &H22, &HCF, &HB4, &HCC, &H38, &HDB, &H50
    
    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetVRInstance
    End If

End Function

Public Function GetEVRInstance() As IBaseFilter

    '&Hfa10746c, &H9b63, &H4b6c, &Hbc, &H49, &Hfc, &H30, &H0e, &Ha5, &Hf2, &H56
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    DEFINE_GUID uuidSrc, &HFA10746C, &H9B63, &H4B6C, &HBC, &H49, &HFC, &H30, &HE, &HA5, &HF2, &H56
    
    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then
        objClassFactory.CreateInstance Nothing, getUnknownUUID, GetVRInstance
    End If

End Function

Private Function GetIGraphBuilder(Optional objFilManager As FilgraphManager = Nothing) As IGraphBuilder

    If (objFilManager Is Nothing) Then
        Set GetIGraphBuilder = objGlobalFilManager
    Else
        Set GetIGraphBuilder = objFilManager
    End If

End Function

Public Function RaiseRender(enumRenderType As RenderType) As IBaseFilter

    Select Case enumRenderType
    
        Case RenderType.MadVRednerer
            Set RaiseRender = GetMadVRInstance
            
        Case RenderType.VideoMixedRenderer9
            Set RaiseRender = GetVMR9Instance
        
        Case RenderType.VideoMixedRenderer
            Set RaiseRender = GetVMR7Instance
            
        Case RenderType.EnhancedVideoRenderer
            Set RaiseRender = GetEVRInstance
            
        Case RenderType.VideoRenderer
            Set RaiseRender = GetVRInstance
    End Select

End Function

Public Function BuildGraph(ByVal strMediaFile As String, _
                           ByRef objFilGraph As FilgraphManager, _
                           ByRef hasVideo As Boolean, _
                           ByRef hasAudio As Boolean, _
                           ByRef hasSubtitle As Boolean, _
                           Optional ByRef Renderer As RenderType = MadVRednerer) As FilgraphManager
    Set objGlobalFilManager = New FilgraphManager
    
    Dim objGraph       As IGraphBuilder

    Dim objSplitter    As IBaseFilter, objAudio As IBaseFilter, objVideo As IBaseFilter, objSubtitle As IBaseFilter, objRender As IBaseFilter

    Dim objSrc         As IBaseFilter
    
    Dim obj_OUT_SrcPin As IPin, obj_IN_Splitter As IPin, objSplitterEnums As IEnumPins

    Dim obj_OUT_Splitter_Audio As IPin
    
    Set objGraph = objGlobalFilManager
    Set objSplitter = GetLAVSplitterInstance
    
    With objGraph
        .AddSourceFilter strMediaFile, "Source", objSrc
        .AddFilter objSplitter, "Splitter"
    
        objSrc.FindPin "Output", obj_OUT_SrcPin
        objSplitter.FindPin "Input", obj_IN_Splitter
        obj_OUT_SrcPin.ConnectedTo obj_IN_Splitter
        
        objSplitter.EnumPins objSplitterEnums

        Dim lngCount   As Long

        Dim lngCurrent As Long

        Dim objCurrPin As IPin

        lngCurrent = 1
        objSplitterEnums.Next lngCurrent, objCurrPin, lngCount

        While (objSplitterEnums.Next(lngCurrent, objCurrPin, lngCount) <> 0)

            Dim sPinInfo As [_PinInfo], sPinName As String

            objCurrPin.QueryPinInfo sPinInfo
            sPinName = AllocStr(sPinInfo.achName(0))

            If (sPinName = "Video") Then
                'Exist Video
                
                Set objVideo = GetLAVVideoInstance
                Set objRender = RaiseRender(Renderer)
                .AddFilter objVideo, "Video"
                .AddFilter objRender, "Render"
                hasVideo = True
                
            ElseIf (sPinName = "Audio") Then
                Set objAudio = GetLAVAudioInstance
                .AddFilter objAudio, "Audio"
                hasAudio = True
                
            ElseIf (sPinName = "Subtitle") Then
                hasSubtitle = True
                
            End If
            
        Wend

        Set objSubtitle = GetVSFilterInstance
        .AddFilter objSubtitle, "Subtitle"
        objSrc.Run
    End With
    
End Function
