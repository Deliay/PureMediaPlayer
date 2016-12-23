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

Private objSplitter             As IBaseFilter, objAudio As IBaseFilter, objVideo As IBaseFilter, objSubtitle As IBaseFilter, objRender As IBaseFilter

Public EVRFilterStorage   As IBaseFilter

Private Declare Function DispCallFunc& _
                Lib "oleaut32" (ByVal ppv&, _
                                ByVal oVft&, _
                                ByVal CC As Long, _
                                ByVal rtTYP%, _
                                ByVal paCount&, _
                                paTypes%, _
                                paValues&, _
                                fuReturn)

Private Declare Function OleCreatePropertyFrame& _
                Lib "oleaut32" (ByVal hwndOwner&, _
                                ByVal X&, _
                                ByVal Y&, _
                                ByVal lpszCaption&, _
                                ByVal cObjects&, _
                                ByRef ppUnk&, _
                                ByVal cPages&, _
                                ByVal pPageClsID&, _
                                ByVal lcid&, _
                                ByVal dwReserved&, _
                                ByVal pvReserved&)

Private Const CLSID_ActiveMovieCategories = "{DA4E3DA0-D07D-11d0-BD50-00A0C911CE86}"

Private Const CLSID_VideoInputDeviceCategory = "{860BB310-5D01-11d0-BD3B-00A0C911CE86}"

Private Function getClassFactoryUUID() As olelib.UUID

    If (uuidIClassFactoryInited = False) Then
        '00000001-0000-0000-C000-000000000046
        'DEFINE_GUID uuidIClassFactory, &H1, &H0, &H0, &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46
        CLSIDFromString "{00000001-0000-0000-C000-000000000046}", uuidIClassFactory
        uuidIClassFactoryInited = True
    End If

    getClassFactoryUUID = uuidIClassFactory
    
End Function

Private Function getUnknownUUID() As olelib.UUID

    If (uuidIUnknownInited = False) Then
        '00000000-0000-0000-C000-000000000046
        'DEFINE_GUID uuidIUnknown, &H0, &H0, &H0, &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46
        CLSIDFromString "{00000000-0000-0000-C000-000000000046}", uuidIUnknown
        uuidIUnknownInited = True
    End If

    getUnknownUUID = uuidIUnknown
End Function

Private Function DEFINE_GUID(ByRef ruuid As olelib.UUID, _
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
                            ByVal k As Byte)

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

Public Sub GetLAVAudioInstance(ByRef instance As IBaseFilter)

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    '{E8E73B6B-4CB3-44A4-BE99-4F7BCB96E491}
    'DEFINE_GUID uuidSrc, &HE8E73B6B, &H4CB3, &H44A4, &HBE, &H99, &H4F, &H7B, &HCB, &H96, &HE4, &H91
    CLSIDFromString "{E8E73B6B-4CB3-44A4-BE99-4F7BCB96E491}", uuidSrc

    If (LAVAudioInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If
    
End Sub

Public Sub GetLAVVideoInstance(ByRef instance As IBaseFilter)

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'EE30215D-164F-4A92-A4EB-9D4C13390F9F
    CLSIDFromString "{EE30215D-164F-4A92-A4EB-9D4C13390F9F}", uuidSrc
    'DEFINE_GUID uuidSrc, &HEE30215D, &H164F, &H4A92, &HA4, &HEB, &H9D, &H4C, &H13, &H39, &HF, &H9F
    
    If (LAVVideoInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If
    
End Sub

Public Sub GetLAVSplitterInstance(ByRef instance As IBaseFilter)

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'B98D13E7-55DB-4385-A33D-09FD1BA26338
    '171252A0-8820-4AFE-9DF8-5C92B2D66B04
    'DEFINE_GUID uuidSrc, &H171252A0, &H8820, &H4AFE, &H9D, &HF8, &H5C, &H92, &HB2, &HD6, &H6B, &H4
    'DEFINE_GUID uuidSrc, &HB98D13E7, &H55DB, &H4385, &HA3, &H3D, &H9, &HFD, &H1B, &HA2, &H63, &H38
    CLSIDFromString "{B98D13E7-55DB-4385-A33D-09FD1BA26338}", uuidSrc

    If (LAVSplitterInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If
    
End Sub

Public Sub GetVSFilterInstance(ByRef instance As IBaseFilter)

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    '93A22E7A-5091-45ef-BA61-6DA26156A5D0
    CLSIDFromString "{93A22E7A-5091-45ef-BA61-6DA26156A5D0}", uuidSrc
    'DEFINE_GUID uuidSrc, &H93A22E7A, &H5091, &H45EF, &HBA, &H61, &H6D, &HA2, &H61, &H56, &HA5, &HD0
    
    If (VSFilterInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If
    
End Sub

Public Sub GetMadVRInstance(ByRef instance As IBaseFilter)

    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    CLSIDFromString "{e1a8b82a-32ce-4b0d-be0d-aa68c772e423}", uuidSrc
    'DEFINE_GUID uuidSrc, &HE1A8B82A, &H32CE, &H4B0D, &H4B, &HD, &HAA, &H68, &HC7, &H72, &HE4, &H23
    
    If (MadVRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If
    
End Sub

Public Sub GetVMR9Instance(ByRef instance As IBaseFilter)

    ' &H51b4abf3, &H748f, &H4e3b, &Ha2, &H76, &Hc8, &H28, &H33, &H0e, &H92, &H6a
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory
    
    DEFINE_GUID uuidSrc, &H51B4ABF3, &H748F, &H4E3B, &HA2, &H76, &HC8, &H28, &H33, &HE, &H92, &H6A

    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If
    
End Sub

Public Sub GetVMR7Instance(ByRef instance As IBaseFilter)

    '&Hb87beb7b, &H8d29, &H423f, &Hae, &H4d, &H65, &H82, &Hc1, &H01, &H75, &Hac
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    DEFINE_GUID uuidSrc, &HB87BEB7B, &H8D29, &H423F, &HAE, &H4D, &H65, &H82, &HC1, &H1, &H75, &HAC
    
    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If

End Sub

Public Sub GetVRInstance(ByRef instance As IBaseFilter)

    '&H6bc1cffa, &H8fc1, &H4261, &Hac, &H22, &Hcf, &Hb4, &Hcc, &H38, &Hdb, &H50
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    'DEFINE_GUID uuidSrc, &H6BC1CFFA, &H8FC1, &H4261, &HAC, &H22, &HCF, &HB4, &HCC, &H38, &HDB, &H50
    CLSIDFromString "{e1a8b82a-32ce-4b0d-be0d-aa68c772e423}", uuidSrc

    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If

End Sub

Public Sub GetEVRInstance(ByRef instance As IBaseFilter)

    '&Hfa10746c, &H9b63, &H4b6c, &Hbc, &H49, &Hfc, &H30, &H0e, &Ha5, &Hf2, &H56
    Dim uuidSrc As olelib.UUID, objClassFactory As IClassFactory

    'e1a8b82a-32ce-4b0d-be0d-aa68c772e423
    DEFINE_GUID uuidSrc, &HFA10746C, &H9B63, &H4B6C, &HBC, &H49, &HFC, &H30, &HE, &HA5, &HF2, &H56
    
    If (VRInstance(uuidSrc, getClassFactoryUUID, objClassFactory) = 0) Then

        Dim objUnknown As olelib.IUnknown

        objClassFactory.CreateInstance Nothing, getUnknownUUID, objUnknown
        Set instance = objUnknown
    End If

End Sub

Private Function GetIGraphBuilder(Optional objFilManager As FilgraphManager = Nothing) As IGraphBuilder

    If (objFilManager Is Nothing) Then
        Set GetIGraphBuilder = objGlobalFilManager
    Else
        Set GetIGraphBuilder = objFilManager
    End If

End Function

Public Sub RaiseRender(enumRenderType As RenderType, ByRef instance As IBaseFilter)

    Select Case enumRenderType
    
        Case RenderType.MadVRednerer
            GetMadVRInstance instance
            
        Case RenderType.VideoMixedRenderer9
            GetVMR9Instance instance
        
        Case RenderType.VideoMixedRenderer
            GetVMR7Instance instance
            
        Case RenderType.EnhancedVideoRenderer
            GetEVRInstance instance
            
        Case RenderType.VideoRenderer
            GetVRInstance instance
    End Select

End Sub

Public Sub BuildGraph(ByVal strMediaFile As String, _
                      ByRef objFilGraph As FilgraphManager, _
                      ByRef hasVideo As Boolean, _
                      ByRef hasAudio As Boolean, _
                      ByRef hasSubtitle As Boolean, _
                      Optional ByRef Renderer As RenderType = MadVRednerer)

    Dim objGraph       As IGraphBuilder

    Dim objSrc         As IBaseFilter
    
    Dim obj_OUT_SrcPin As IPin, obj_IN_Splitter As IPin, objSplitterEnums As IEnumPins

    Dim obj_OUT_Audio  As IPin, obj_OUT_Video As IPin, obj_OUT_Subtitle As IPin
    
    Set objFilGraph = New FilgraphManager
    Set objGraph = objFilGraph
    
    GetLAVSplitterInstance objSplitter
    
    With objGraph

        .AddFilter objSplitter, "Splitter"
        
        Dim objSplitterInfo As IFilterInfo
        objFilGraph.FilterCollection.Item 0, objSplitterInfo
        objSplitterInfo.FileName = strMediaFile
 
        objSplitter.EnumPins objSplitterEnums
    
        Dim lngCount   As Long
    
        Dim lngCurrent As Long
    
        Dim objCurrPin As IPin
    
        Dim sPinInfo As [_PinInfo], sPinName As String
    
        lngCurrent = 1

        While (objSplitterEnums.Next(lngCurrent, objCurrPin, lngCount) = 0)

            objCurrPin.QueryPinInfo sPinInfo
            sPinName = AllocStr(sPinInfo.achName(0))
    
            If (sPinName = "Video") Then
                Set obj_OUT_Video = objCurrPin
                hasVideo = True
                
            ElseIf (sPinName = "Audio") Then
                Set obj_OUT_Audio = objCurrPin
                hasAudio = True
                
            ElseIf (sPinName = "Subtitle") Then
                Set obj_OUT_Subtitle = objCurrPin
                hasSubtitle = True
                
            End If
            
        Wend

        Dim objAudioPinInput As IPin, objAudioPinOutput As IPin

        Dim objVideoPinInput As IPin, objVideoPinOutput As IPin

        Dim objSubPinInput   As IPin, objSubPinVideo As IPin, objSubPinOutput As IPin
        
        Dim objRenderPinInput As IPin

        If (hasAudio) Then
            GetLAVAudioInstance objAudio
            .AddFilter objAudio, "Audio"
            
            objAudio.EnumPins objSplitterEnums
            lngCurrent = 1

            While (objSplitterEnums.Next(lngCurrent, objCurrPin, lngCount) = 0)

                objCurrPin.QueryPinInfo sPinInfo
                sPinName = AllocStr(sPinInfo.achName(0))

                If (sPinName = "Output") Then
                    Set objAudioPinOutput = objCurrPin
                ElseIf (sPinName = "Input") Then
                    Set objAudioPinInput = objCurrPin
                End If

            Wend
            
            .Connect obj_OUT_Audio, objAudioPinInput
            .Render objAudioPinOutput
        End If

        If (hasVideo) Then
            GetLAVVideoInstance objVideo
            RaiseRender Renderer, objRender
            GetVSFilterInstance objSubtitle

            .AddFilter objVideo, "Video"
            .AddFilter objRender, "Render"
            .AddFilter objSubtitle, "Subtitle"

            objSubtitle.EnumPins objSplitterEnums
            lngCurrent = 1

            While (objSplitterEnums.Next(lngCurrent, objCurrPin, lngCount) = 0)

                objCurrPin.QueryPinInfo sPinInfo
                sPinName = AllocStr(sPinInfo.achName(0))

                If (sPinName = "Video") Then
                    Set objSubPinVideo = objCurrPin
                ElseIf (sPinName = "Output") Then
                    Set objSubPinOutput = objCurrPin
                ElseIf (sPinName = "Input") Then
                    Set objSubPinInput = objCurrPin
                End If

            Wend

            If (hasSubtitle) Then
                .Connect obj_OUT_Subtitle, objSubPinInput
            End If

            objVideo.EnumPins objSplitterEnums
            lngCurrent = 1

            While (objSplitterEnums.Next(lngCurrent, objCurrPin, lngCount) = 0)

                objCurrPin.QueryPinInfo sPinInfo
                sPinName = AllocStr(sPinInfo.achName(0))

                If (sPinName = "Input") Then
                    Set objVideoPinInput = objCurrPin
                ElseIf (sPinName = "Output") Then
                    Set objVideoPinOutput = objCurrPin
                End If

            Wend

            .Connect obj_OUT_Video, objVideoPinInput
            .Connect objVideoPinOutput, objSubPinVideo
            objRender.EnumPins objSplitterEnums
            lngCurrent = 1
            objSplitterEnums.Next lngCurrent, objRenderPinInput, lngCount
            objRenderPinInput.QueryPinInfo sPinInfo
            sPinName = AllocStr(sPinInfo.achName(0))
            .Connect objSubPinOutput, objRenderPinInput
        End If

        
    End With

End Sub


Private Function CastToIUnknow(ByVal Flt As olelib.IUnknown) As olelib.IUnknown
    Set CastToIUnknow = Flt

End Function

Public Function SetVSFilterFileName(FileName As String) As Boolean

    Const IID_IDirectVobSub = "{EBE1FB08-3957-47ca-AF13-5827E5442E56}", VTbl_SetFileName = 4

    Dim oDirectVobSub As stdole.IUnknown

    Set oDirectVobSub = CastToUnkByIID(objSubtitle, IID_IDirectVobSub)
    SetVSFilterFileName = vtblCall(ObjPtr(oDirectVobSub), VTbl_SetFileName, vbEmpty, StrPtr(FileName)) = S_OK
    
    If (GlobalConfig.SubtitleBind.Exist(mdlGlobalPlayer.FileMD5)) Then

        GlobalConfig.SubtitleBind.Value(mdlGlobalPlayer.FileMD5) = FileName

    Else

        GlobalConfig.SubtitleBind.AddKeyValue mdlGlobalPlayer.FileMD5, FileName

    End If

End Function

Public Function GetVSFilterLangCount() As Long
    Const IID_IDirectVobSub = "{EBE1FB08-3957-47ca-AF13-5827E5442E56}", VTbl_get_LanguageCount = 5

    Dim oDirectVobSub As stdole.IUnknown

    Set oDirectVobSub = CastToUnkByIID(objSubtitle, IID_IDirectVobSub)
    vtblCall ObjPtr(oDirectVobSub), VTbl_get_LanguageCount, vbEmpty, VarPtr(GetVSFilterLangCount)

End Function

Public Function GetVSFilterLangName(ByVal iLanguage As Long) As String
    Const IID_IDirectVobSub = "{EBE1FB08-3957-47ca-AF13-5827E5442E56}", VTbl_get_LanguageName = 6

    Dim oDirectVobSub As stdole.IUnknown

    Set oDirectVobSub = CastToUnkByIID(objSubtitle, IID_IDirectVobSub)
    
    Dim stringbuf As String
    stringbuf = Space(256)
    vtblCall ObjPtr(oDirectVobSub), VTbl_get_LanguageName, vbEmpty, iLanguage, StrPtr(stringbuf)
    GetVSFilterLangName = stringbuf
End Function

Private Function CheckForFileSinkAndSetFileName(ByVal Flt As olelib.IUnknown, _
                                                FileName As String) As Boolean

    Const IID_IFileSinkFilter = "{A2104830-7C70-11CF-8BCE-00AA00A3F1A6}", VTbl_SetFileName = 3

    Dim oUnkFSink As stdole.IUnknown

    Set oUnkFSink = CastToUnkByIID(Flt, IID_IFileSinkFilter)
    CheckForFileSinkAndSetFileName = vtblCall(ObjPtr(oUnkFSink), VTbl_SetFileName, vbLong, StrPtr(FileName), 0&) = S_OK

End Function

Private Function CastToUnkByIID(ByVal ObjToCastFrom As olelib.IUnknown, _
                                IID As String) As stdole.IUnknown

    Dim UUID As olelib.UUID

    olelib.CLSIDFromString IID, UUID
    ObjToCastFrom.QueryInterface UUID, CastToUnkByIID

End Function

Private Function vtblCall(ByVal pUnk As Long, _
                          ByVal vtblIdx As Long, _
                          ByVal retType As VbVarType, _
                          ParamArray P() As Variant)

    Static VType(0 To 31) As Integer, VPtr(0 To 31) As Long

    Dim i As Long, v(), HResDisp As Long

    If pUnk = 0 Then vtblCall = 5: Exit Function

    v = P 'make a copy of the params, to prevent problems with VT_ByRef-Members in the ParamArray

    For i = 0 To UBound(v)
        VType(i) = VarType(v(i))
        VPtr(i) = VarPtr(v(i))
    Next i
    
    HResDisp = DispCallFunc(pUnk, vtblIdx * 4, 4, retType, i, VType(0), VPtr(0), vtblCall)

    If HResDisp <> S_OK Then Err.Raise HResDisp

End Function

Public Function ShowVideoDecoderConfig()

    If (IsEmpty(objVideo) Or objVideo Is Nothing) Then
        
        GetLAVVideoInstance objVideo

    End If

    ShowPropertyPage objVideo, "PureMediaPlayer - Video", frmMain.hWnd

End Function

Public Function ShowAudioDecoderConfig()

    If (IsEmpty(objAudio) Or objAudio Is Nothing) Then
        GetLAVAudioInstance objAudio

    End If

    ShowPropertyPage objAudio, "PureMediaPlayer - Audio", frmMain.hWnd

End Function

Public Function ShowSpliterConfig()

    If (IsEmpty(objSplitter) Or objSplitter Is Nothing) Then
        GetLAVSplitterInstance objSplitter

    End If

    ShowPropertyPage objSplitter, "PureMediaPlayer - Splitter", frmMain.hWnd

End Function

Public Function ShowSubtitleConfig()

    If (IsEmpty(objSubtitle) Or objSubtitle Is Nothing) Then
        GetVSFilterInstance objSubtitle

    End If

    ShowPropertyPage objSubtitle, "PureMediaPlayer - Subtitle", frmMain.hWnd

End Function

Public Function ShowRendererConfig()

    If (IsEmpty(objRender) Or objRender Is Nothing) Then
    
        RaiseRender GlobalRenderType, objRender

    End If
    
    ShowPropertyPage objRender, "PureMediaPlayer - Renderer", frmMain.hWnd

End Function

Public Function ShowPropertyPage(ByVal FilterOrPin As olelib.IUnknown, _
                                 Optional Caption As String, _
                                 Optional ByVal hwndOwner As Long) As Boolean

    Const IID_ISpecifyPropertyPages = "{B196B28B-BAB4-101A-B69C-00AA00341D07}", VTbl_GetPages = 3

    Dim oUnkSpPP As stdole.IUnknown, CAUUID(0 To 1) As Long

    Set oUnkSpPP = CastToUnkByIID(FilterOrPin, IID_ISpecifyPropertyPages)

    If vtblCall(ObjPtr(oUnkSpPP), vbLong, VTbl_GetPages, VarPtr(CAUUID(0))) Then Exit Function
    If CAUUID(0) = 0 Then Exit Function 'no PropPageCount was returned

    OleCreatePropertyFrame hwndOwner, 0, 0, StrPtr(Caption), 1, ObjPtr(FilterOrPin), CAUUID(0), CAUUID(1), 0, 0, 0

    CoTaskMemFree CAUUID(1)
    ShowPropertyPage = True

End Function

Public Sub QueryMediaStreams()

    Dim objStreams As Strmif.IAMStreamSelect
    Dim lngCount As Long, lngIter As Long, lngFlag As Long, lngGroup As Long, strName As String, strCpyName As String
    Dim lngCurrGroup As Long
    Set objStreams = objSplitter
    strName = Space$(255)
    
    objStreams.Count lngCount
    For lngIter = 0 To lngCount - 1

        objStreams.Info lngIter, 0#, lngFlag, 0, lngGroup, strName, 0#, 0#
        If (lngIter > frmMenu.mmStatus_Streams_Select.Count - 1) Then Load frmMenu.mmStatus_Streams_Select(lngIter)
        strCpyName = Mid(strName, 1, InStr(1, strName, Chr(0)))
        frmMenu.mmStatus_Streams_Select(lngIter).Caption = CStr(strCpyName)
        If (lngFlag <> 0) Then
            frmMenu.mmStatus_Streams_Select(lngIter).Checked = True
            
        Else
            frmMenu.mmStatus_Streams_Select(lngIter).Checked = False

        End If
        CoTaskMemFree StrPtr(strName)
        
    Next
    
    If (frmMenu.mmStatus_Streams_Select.Count > lngCount) Then
        For lngIter = lngCount To frmMenu.mmStatus_Streams_Select.Count
            frmMenu.mmStatus_Streams_Select(lngIter).Visible = False

        Next
        
    End If
    
End Sub

Public Sub EnableMediaStreams(ByVal enableIndex As Long)

    Dim objStreams As Strmif.IAMStreamSelect
    Set objStreams = objSplitter
    
    objStreams.Enable enableIndex, ByVal 3
    
End Sub
