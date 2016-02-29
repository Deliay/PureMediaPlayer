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

''
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

Private LAVVideoIndex    As Long, LAVAudioIndex As Long

Private LAVSplitterIndex As Long, LAVSplitterSourceIndex As Long

Private VSFilterIndex    As Long, EVRIndex As Long, MadVRIndex As Long

Private VMR9Index        As Long, VMR7Index As Long, VRIndex As Long

Public EVRFilterStorage  As IBaseFilter

Public Sub RegisterAllDecoder()
    RegisterLAVAudio
    RegisterLAVSplitter
    RegisterLAVVideo
    RegisterVSFilter
    RegisterMadVRFilter
    
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
    mdlToolBarAlphaer.LoadUI
    
End Sub

Public Sub BuildGrph(ByVal srcFile As String, _
                     ByRef objGraphManager As FilgraphManager, _
                     ByRef boolHasVideo As Boolean, _
                     ByRef boolHasAudio As Boolean, _
                     ByRef boolHasSubtitle As Boolean, _
                     Optional ByVal eRenderer As RenderType = VideoMixedRenderer9)

    Dim objSrcFileFilter  As IFilterInfo

    Dim objSrcSplitterReg As IRegFilterInfo, objSrcSplitterFilter As IFilterInfo

    Dim objVideoReg       As IRegFilterInfo, objVideoFilter As IFilterInfo, objVideoPin As IPinInfo

    Dim objAudioReg       As IRegFilterInfo, objAudioFilter As IFilterInfo, objAudioPin As IPinInfo

    Dim objSubtitleReg    As IRegFilterInfo, objSubtitleFilter As IFilterInfo, objSubtitlePin As IPinInfo

    Dim objRenderReg      As IRegFilterInfo, objRenderFilter As IFilterInfo ', objRenderPin As IPinInfo

    Dim objSpliterPin     As IPinInfo

    'Try put source file directly
    '
    'On Error GoTo OnlyInput
    'objGraphManager.AddSourceFilter srcFile, objSrcSplitterFilter

    '    If (Not objSrcSplitterFilter Is Nothing And objSrcSplitterFilter.Pins.Count > 0) Then
    '        GoTo ParserPins
    'OnlyInput:
    '    'Create Splitter
    '        Resume Next
    '    Else
    objGraphManager.RegFilterCollection.Item LAVSplitterSourceIndex, objSrcSplitterReg
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
    
    If (Dir(srcFile) <> "") Then
        objGraphManager.AddSourceFilter srcFile, objSrcSplitterFilter

        For Each objSpliterPin In objSrcSplitterFilter.Pins

            objSpliterPin.Render
            boolHasAudio = True
        Next

    End If

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

Public Function CastToUnkByIID(ByVal ObjToCastFrom As olelib.IUnknown, _
                               IID As String) As stdole.IUnknown

    Dim UUID As olelib.UUID

    olelib.CLSIDFromString IID, UUID
    ObjToCastFrom.QueryInterface UUID, CastToUnkByIID

End Function

Public Function vtblCall(ByVal pUnk As Long, _
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
