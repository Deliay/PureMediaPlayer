Attribute VB_Name = "mdlFilterBuilder"
Option Explicit

Private Declare Function RegisterLAVAudio _
                Lib "LAVAudio.ax" _
                Alias "DllRegisterServer" () As Long

Private Declare Function RegisterLAVSplitter _
                Lib "LAVSplitter.ax" _
                Alias "DllRegisterServer" () As Long

Private Declare Function RegisterLAVVideo _
                Lib "LAVVideo.ax" _
                Alias "DllRegisterServer" () As Long

Private Declare Function RegisterVSFilter _
                Lib "vsfilter.dll" _
                Alias "DllRegisterServer" () As Long

Private Declare Function RegisterMadVRFilter _
                Lib "madVR.ax" _
                Alias "DllRegisterServer" () As Long
                
Private Declare Function RegisterSSubTmr6 _
                Lib "SSubTmr6.dll" _
                Alias "DllRegisterServer" () As Long
                
Private Declare Function RegistervbaListView6 _
                Lib "vbaListView6.ocx" _
                Alias "DllRegisterServer" () As Long

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

Private objSrcSplitterReg As IRegFilterInfo, objSrcSplitterFilter As IFilterInfo

Private objVideoReg       As IRegFilterInfo, objVideoFilter As IFilterInfo, objVideoPin As IPinInfo

Private objAudioReg       As IRegFilterInfo, objAudioFilter As IFilterInfo, objAudioPin As IPinInfo

Private objSubtitleReg    As IRegFilterInfo, objSubtitleFilter As IFilterInfo, objSubtitlePin As IPinInfo

Private objRenderReg      As IRegFilterInfo, objRenderFilter As IFilterInfo ', objRenderPin As IPinInfo

Private objSpliterPin     As IPinInfo

Private LAVVideoIndex    As Long, LAVAudioIndex As Long

Private LAVSplitterIndex As Long, LAVSplitterSourceIndex As Long

Private VSFilterIndex    As Long, EVRIndex As Long, MadVRIndex As Long

Private VMR9Index        As Long, VMR7Index As Long, VRIndex As Long

Public EVRFilterStorage  As IBaseFilter
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub RegisterCOM()
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\SSubTmr6.dll", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\vbaListView6.ocx", App.Path & "\", 0
    RegisterSSubTmr6
    RegistervbaListView6
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\SSubTmr6.dll", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\vbaListView6.ocx", App.Path & "\", 0
    
End Sub

Public Sub RegisterAllDecoder()
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\" & "LAVAudio.ax", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\" & "LAVSplitter.ax", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\" & "LAVVideo.ax", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\" & "madVR.ax", App.Path & "\", 0
    RegisterLAVAudio
    RegisterLAVSplitter
    RegisterLAVVideo
    RegisterVSFilter
    RegisterMadVRFilter
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\" & "LAVAudio.ax", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\" & "LAVSplitter.ax", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\" & "LAVVideo.ax", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\" & "madVR.ax", App.Path & "\", 0
End Sub

Public Sub Main()
    LAVVideoIndex = -1
    LAVAudioIndex = -1
    LAVSplitterIndex = -1
    LAVSplitterSourceIndex = -1
    VSFilterIndex = -1
    VMR9Index = -1
    EVRIndex = -1
    MadVRIndex = -1
    FillDecoder mdlGlobalPlayer.GlobalFilGraph
    
    If (LAVAudioIndex = -1 Or LAVVideoIndex = -1 Or LAVSplitterIndex = -1 Or VSFilterIndex = -1 Or LAVSplitterSourceIndex = -1 Or MadVRIndex = -1) Then
        
        RegisterAllDecoder
        Set mdlGlobalPlayer.GlobalFilGraph = New FilgraphManager
        FillDecoder mdlGlobalPlayer.GlobalFilGraph

    End If
    
    If (VMR9Index = -1 And EVRIndex = -1 And VMR7Index = -1) Then
        MsgBox "Bad Renderer Support!"
        End

    End If
    
    If (LAVAudioIndex = -1 Or LAVVideoIndex = -1 Or LAVSplitterIndex = -1 Or VSFilterIndex = -1 Or LAVSplitterSourceIndex = -1 Or MadVRIndex = -1) Then
    
        MsgBox "Cannot Register Decoder! Please run again with Administrator Permission"
        End

    End If

    Load frmMenu

    If (VMR9Index = -1) Then
        frmMenu.Renderers(RenderType.VideoMixedRenderer9).Enabled = False

    End If

    If (EVRIndex = -1) Then
        frmMenu.Renderers(RenderType.EnhancedVideoRenderer).Enabled = False

    End If

    If (getConfig("Renderer") = "") Then
        frmMenu.Renderers_Click RenderType.MadVRednerer
    Else
        frmMenu.Renderers_Click val(getConfig("Renderer"))

    End If

    'Create A Clean FilgraphManager
    Set mdlGlobalPlayer.GlobalFilGraph = New FilgraphManager
    
    On Error GoTo RegisterCOMErr
    mdlToolBarAlphaer.LoadUI
    On Error GoTo 0
    Exit Sub
RegisterCOMErr:
    'do gui com register
    RegisterCOM
End Sub

Public Sub BuildGrph(ByVal srcFile As String, _
                     ByRef objGraphManager As FilgraphManager, _
                     ByRef boolHasVideo As Boolean, _
                     ByRef boolHasAudio As Boolean, _
                     ByRef boolHasSubtitle As Boolean, _
                     Optional ByVal eRenderer As RenderType = VideoMixedRenderer9)



    'Try put source file directly
    '
    'On Error GoTo OnlyInput
    'objGraphManager.AddSourceFilter srcFile, objSrcSplitterFilter

    'If (Not objSrcSplitterFilter Is Nothing And objSrcSplitterFilter.Pins.Count > 0) Then GoTo ParserPins

    objGraphManager.RegFilterCollection.Item LAVSplitterSourceIndex, objSrcSplitterReg
    On Error GoTo regControl
    objSrcSplitterReg.Filter objSrcSplitterFilter

    On Error GoTo notExist

    objSrcSplitterFilter.FileName = srcFile
    'Add Src File
    CheckForFileSinkAndSetFileName objSrcSplitterFilter, srcFile
    '    End If
    'Reset switch
    boolHasVideo = False
    boolHasAudio = False
    boolHasSubtitle = False
ParserPins:

    On Error GoTo 0

    'Create Filters
    For Each objSpliterPin In objSrcSplitterFilter.Pins

        If (objSpliterPin.Name = "Audio") Then
            boolHasAudio = True
            objGraphManager.RegFilterCollection.Item LAVAudioIndex, objAudioReg
            objAudioReg.Filter objAudioFilter
            
            Set objAudioPin = objSpliterPin
            
        ElseIf (objSpliterPin.Name = "Video") Then
            boolHasVideo = True
            objGraphManager.RegFilterCollection.Item LAVVideoIndex, objVideoReg
            objVideoReg.Filter objVideoFilter

            Dim VMRVerSpec As Long

            If (eRenderer = VideoMixedRenderer9) Then
                VMRVerSpec = VMR9Index
                
            ElseIf (eRenderer = VideoMixedRenderer) Then
                VMRVerSpec = VMR7Index
                
            ElseIf (eRenderer = EnhancedVideoRenderer) Then
                VMRVerSpec = EVRIndex
                
            ElseIf (eRenderer = MadVRednerer) Then
                VMRVerSpec = MadVRIndex
                
            Else
                VMRVerSpec = VRIndex

            End If
            
            objGraphManager.RegFilterCollection.Item VMRVerSpec, objRenderReg
            objRenderReg.Filter objRenderFilter
                
            If (eRenderer = EnhancedVideoRenderer) Then

                Dim GraphBuilder As IGraphBuilder: Set GraphBuilder = objGraphManager

                GraphBuilder.FindFilterByName "Enhanced Video Renderer", EVRFilterStorage

            End If
            
            Set objVideoPin = objSpliterPin
            
        ElseIf (objSpliterPin.Name = "Subtitle") Then
            boolHasSubtitle = True
            objGraphManager.RegFilterCollection.Item VSFilterIndex, objSubtitleReg
            objSubtitleReg.Filter objSubtitleFilter
            
            Set objSubtitlePin = objSpliterPin
            
        End If

    Next
    
    '1. LAVSplit (Subtitle) -> (Input) VSFilter
    '2. LAVSplit (Video) -> (Input) LAVVideo (Output) -> (Video) VSFilter (Output)-> (VMR Input)(VMR9)
    '3. LAVSplit (Audio) -> (Input) LAVAudio (Output) -> Connect Downstream
    
    '  Audio first
    If (boolHasAudio = True) Then objAudioPin.Render 'audio auto connect

    '  Video second
    If (boolHasVideo = True) Then
        '  If subtitle exist
        objVideoPin.Render

        ' LAVSplit (Subtitle) -> (Input) VSFilter
        If (boolHasSubtitle = True) Then

            Dim objPinVSInput As IPinInfo

            For Each objPinVSInput In objSubtitleFilter.Pins

                If (LCase(objPinVSInput.Name) = "input") Then Exit For
            Next
            objSubtitlePin.Connect objPinVSInput

        End If

    End If

    Exit Sub
notExist:
    Set GlobalFilGraph = Nothing
Exit Sub
regControl:
    mdlFilterBuilder.RegisterAllDecoder
End Sub

Private Sub FillDecoder(m_GraphManager As FilgraphManager)

    Dim i             As Long

    Dim objRegFilter  As IRegFilterInfo

    Dim objTestFilter As IFilterInfo, objTestPin As IPinInfo

    For Each objRegFilter In m_GraphManager.RegFilterCollection

        i = i + 1

        If (objRegFilter.Name = "LAV Splitter") Then
            LAVSplitterIndex = i - 1
            
        ElseIf (objRegFilter.Name = "LAV Splitter Source") Then
            LAVSplitterSourceIndex = i - 1
            
        ElseIf (objRegFilter.Name = "LAV Video Decoder") Then
            LAVVideoIndex = i - 1
            
        ElseIf (objRegFilter.Name = "LAV Audio Decoder") Then
            LAVAudioIndex = i - 1
            
        ElseIf (objRegFilter.Name = "VSFilter") Then
            VSFilterIndex = i - 1
            
        ElseIf (objRegFilter.Name = "Video Mixing Renderer 9") Then
            VMR9Index = i - 1
            
        ElseIf (objRegFilter.Name = "Video Renderer") Then
            objRegFilter.Filter objTestFilter

            For Each objTestPin In objTestFilter.Pins

                If (LCase$(objTestPin.Name) = "vmr input0") Then
                    VMR7Index = i - 1
                    Exit For
                ElseIf (LCase$(objTestPin.Name) = "input") Then
                    VRIndex = i - 1
                    Exit For

                End If

            Next
            
        ElseIf (objRegFilter.Name = "Enhanced Video Renderer") Then
            EVRIndex = i - 1
            
        ElseIf (objRegFilter.Name = "madVR") Then
            MadVRIndex = i - 1
            
        End If
        
        If (LAVAudioIndex <> -1 And LAVVideoIndex <> -1 And LAVSplitterIndex <> -1 And VSFilterIndex <> -1 And LAVSplitterSourceIndex <> -1 And VMR9Index <> -1 And VMR7Index <> -1 And VRIndex <> -1 And EVRIndex <> -1 And MadVRIndex <> -1) Then
        
            Exit For
        
        End If

    Next

End Sub

Private Function CheckForFileSinkAndSetFileName(ByVal Flt As olelib.IUnknown, _
                                                FileName As String) As Boolean

    Const IID_IFileSinkFilter = "{A2104830-7C70-11CF-8BCE-00AA00A3F1A6}", VTbl_SetFileName = 3

    Dim oUnkFSink As stdole.IUnknown

    Set oUnkFSink = CastToUnkByIID(Flt, IID_IFileSinkFilter)
    CheckForFileSinkAndSetFileName = vtblCall(ObjPtr(oUnkFSink), VTbl_SetFileName, StrPtr(FileName), 0&) = S_OK

End Function

Private Function CastToUnkByIID(ByVal ObjToCastFrom As olelib.IUnknown, _
                               IID As String) As stdole.IUnknown

    Dim UUID As olelib.UUID

    olelib.CLSIDFromString IID, UUID
    ObjToCastFrom.QueryInterface UUID, CastToUnkByIID

End Function

Private Function vtblCall(ByVal pUnk As Long, _
                         ByVal vtblIdx As Long, _
                         ParamArray P() As Variant)

    Static VType(0 To 31) As Integer, VPtr(0 To 31) As Long

    Dim i As Long, V(), HResDisp As Long

    If pUnk = 0 Then vtblCall = 5: Exit Function

    V = P 'make a copy of the params, to prevent problems with VT_ByRef-Members in the ParamArray

    For i = 0 To UBound(V)
        VType(i) = VarType(V(i))
        VPtr(i) = VarPtr(V(i))
    Next i
    
    HResDisp = DispCallFunc(pUnk, vtblIdx * 4, 4, vbLong, i, VType(0), VPtr(0), vtblCall)

    If HResDisp <> S_OK Then Err.Raise HResDisp

End Function

Public Function ShowVideoDecoderConfig()
    If (HasVideo) Then
        ShowPropertyPage objVideoFilter.Filter, "PureMediaPlayer - " & objVideoFilter.Name, frmMain.hWnd
    Else
        MsgBox "Current dose not have any Video Decoder"
    End If
End Function

Public Function ShowAudioDecoderConfig()
    If (HasAudio) Then
        ShowPropertyPage objAudioFilter.Filter, "PureMediaPlayer - " & objAudioFilter.Name, frmMain.hWnd
    Else
        MsgBox "Current dose not have any Audio Decoder"
    End If
End Function

Public Function ShowSpliterConfig()
    If (mdlGlobalPlayer.Loaded) Then
        ShowPropertyPage objSrcSplitterFilter.Filter, "PureMediaPlayer - " & objSrcSplitterFilter.Name, frmMain.hWnd
    Else
        MsgBox "Current dose not have any Spliterer"
    End If
End Function

Public Function ShowSubtitleConfig()
    If (HasSubtitle) Then
        ShowPropertyPage objSubtitleFilter.Filter, "PureMediaPlayer - " & objSubtitleFilter.Name, frmMain.hWnd
    Else
        MsgBox "Current dose not have any Subtitle"
    End If
End Function

Public Function ShowRendererConfig()
    If (HasVideo) Then
        ShowPropertyPage objRenderFilter.Filter, "PureMediaPlayer - " & objRenderFilter.Name, frmMain.hWnd
    Else
        MsgBox "Current dose not have any Video Renderer"
    End If
End Function

Public Function ShowPropertyPage(ByVal FilterOrPin As olelib.IUnknown, Optional Caption As String, Optional ByVal hwndOwner As Long) As Boolean
    Const IID_ISpecifyPropertyPages = "{B196B28B-BAB4-101A-B69C-00AA00341D07}", VTbl_GetPages = 3
    Dim oUnkSpPP As stdole.IUnknown, CAUUID(0 To 1) As Long
    Set oUnkSpPP = CastToUnkByIID(FilterOrPin, IID_ISpecifyPropertyPages)

    If vtblCall(ObjPtr(oUnkSpPP), VTbl_GetPages, VarPtr(CAUUID(0))) Then Exit Function
    If CAUUID(0) = 0 Then Exit Function 'no PropPageCount was returned

    OleCreatePropertyFrame hwndOwner, 0, 0, StrPtr(Caption), 1, ObjPtr(FilterOrPin), CAUUID(0), CAUUID(1), 0, 0, 0

    CoTaskMemFree CAUUID(1)
    ShowPropertyPage = True
End Function
