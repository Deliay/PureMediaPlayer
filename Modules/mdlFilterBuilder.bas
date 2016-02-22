Attribute VB_Name = "mdlFilterBuilder"
Option Explicit
Private Declare Function RegisterLAVAudio Lib "Filter\LAVAudio.ax" () As Long
Private Declare Function RegisterLAVSplitter Lib "Filter\LAVSplitter.ax" () As Long
Private Declare Function RegisterLAVVideo Lib "Filter\LAVVideo.ax" () As Long
Private Declare Function RegisterVSFilter Lib "Filter\vsfilter.dll" () As Long
''DllRegisterServer
Private Declare Function DispCallFunc& Lib "oleaut32" (ByVal ppv&, ByVal oVft&, ByVal CC As Long, ByVal rtTYP%, ByVal paCount&, paTypes%, paValues&, fuReturn)
Private Declare Function OleCreatePropertyFrame& Lib "oleaut32" (ByVal hWndOwner&, ByVal X&, ByVal Y&, ByVal lpszCaption&, ByVal cObjects&, ByRef ppUnk&, ByVal cPages&, ByVal pPageClsID&, ByVal lcid&, ByVal dwReserved&, ByVal pvReserved&)

Private Const CLSID_ActiveMovieCategories = "{DA4E3DA0-D07D-11d0-BD50-00A0C911CE86}"
Private Const CLSID_VideoInputDeviceCategory = "{860BB310-5D01-11d0-BD3B-00A0C911CE86}"

Private LAVVideoIndex As Long, LAVAudioIndex As Long
Private LAVSplitterIndex As Long, LAVSplitterSourceIndex As Long
Private VSFilterIndex As Long
Private VMR9Index As Long, VMR7Index As Long, VRIndex As Long

Public Sub RegisterAllDecoder()
    RegisterLAVAudio
    RegisterLAVSplitter
    RegisterLAVVideo
    RegisterVSFilter
End Sub

Public Sub Main()
    LAVVideoIndex = -1
    LAVAudioIndex = -1
    LAVSplitterIndex = -1
    LAVSplitterSourceIndex = -1
    VSFilterIndex = -1
    VMR9Index = -1
    Set mdlGlobalPlayer.GlobalFilGraph = New FilgraphManager
    FillDecoder mdlGlobalPlayer.GlobalFilGraph
    
    If (LAVAudioIndex = -1 Or _
        LAVVideoIndex = -1 Or _
        LAVSplitterIndex = -1 Or _
        VSFilterIndex = -1 Or _
        LAVSplitterSourceIndex = -1) Then
        
        RegisterAllDecoder
        Set mdlGlobalPlayer.GlobalFilGraph = New FilgraphManager
        FillDecoder mdlGlobalPlayer.GlobalFilGraph
    End If
    
    If (VMR9Index = -1) Then
        MsgBox "Your computer not support VMR9 Render"
        End
    End If
    
    If (LAVAudioIndex = -1 Or _
        LAVVideoIndex = -1 Or _
        LAVSplitterIndex = -1 Or _
        VSFilterIndex = -1 Or _
        LAVSplitterSourceIndex = -1) Then
    
        MsgBox "Cannot Register Decoder! Please run me with Admin Perm"
        End
    End If
    
    frmMain.Show
    
End Sub

Public Sub BuildGrph(ByVal srcFile As String, _
                          ByRef objGraphManager As FilgraphManager, _
                          ByRef boolHasVideo As Boolean, _
                          ByRef boolHasAudio As Boolean, _
                          ByRef boolHasSubtitle As Boolean)
                          
    Dim objSrcSplitterReg As IRegFilterInfo, objSrcSplitterFilter As IFilterInfo
    Dim objVideoReg As IRegFilterInfo, objVideoFilter As IFilterInfo, objVideoPin As IPinInfo
    Dim objAudioReg As IRegFilterInfo, objAudioFilter As IFilterInfo, objAudioPin As IPinInfo
    Dim objSubtitleReg As IRegFilterInfo, objSubtitleFilter As IFilterInfo, objSubtitlePin As IPinInfo
    Dim objRenderReg As IRegFilterInfo, objRenderFilter As IFilterInfo ', objRenderPin As IPinInfo
    Dim objSpliterPin As IPinInfo
    'Create Splitter
    objGraphManager.RegFilterCollection.Item LAVSplitterSourceIndex, objSrcSplitterReg
    objSrcSplitterReg.Filter objSrcSplitterFilter
    
    objSrcSplitterFilter.FileName = srcFile
    'Add Src File
    CheckForFileSinkAndSetFileName objSrcSplitterFilter, srcFile
    
    'Reset switch
    boolHasVideo = False
    boolHasAudio = False
    boolHasSubtitle = False
    
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
            
            objGraphManager.RegFilterCollection.Item VMR9Index, objRenderReg
            objRenderReg.Filter objRenderFilter
            
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
 
End Sub



Private Sub FillDecoder(m_GraphManager As FilgraphManager)
    Dim i As Long
    Dim objRegFilter As Object
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
        End If
        
        If (LAVAudioIndex <> -1 And _
            LAVVideoIndex <> -1 And _
            LAVSplitterIndex <> -1 And _
            VSFilterIndex <> -1 And _
            LAVSplitterSourceIndex <> -1) Then
        
            Exit For
        
        End If
    Next
End Sub

Private Function CheckForFileSinkAndSetFileName(ByVal Flt As olelib.IUnknown, FileName As String) As Boolean
    Const IID_IFileSinkFilter = "{A2104830-7C70-11CF-8BCE-00AA00A3F1A6}", VTbl_SetFileName = 3
    Dim oUnkFSink As stdole.IUnknown

    Set oUnkFSink = CastToUnkByIID(Flt, IID_IFileSinkFilter)
    CheckForFileSinkAndSetFileName = vtblCall(ObjPtr(oUnkFSink), VTbl_SetFileName, StrPtr(FileName), 0&) = S_OK
End Function

Private Function CastToUnkByIID(ByVal ObjToCastFrom As olelib.IUnknown, IID As String) As stdole.IUnknown
    Dim UUID As olelib.UUID
    olelib.CLSIDFromString IID, UUID
    ObjToCastFrom.QueryInterface UUID, CastToUnkByIID
End Function

Private Function vtblCall(ByVal pUnk As Long, ByVal vtblIdx As Long, ParamArray P() As Variant)
    Static VType(0 To 31) As Integer, VPtr(0 To 31) As Long
    Dim i As Long, V(), HResDisp As Long
    If pUnk = 0 Then vtblCall = 5: Exit Function

    V = P 'make a copy of the params, to prevent problems with VT_ByRef-Members in the ParamArray
    For i = 0 To UBound(V)
        VType(i) = VarType(V(i))
        VPtr(i) = VarPtr(V(i))
    Next i
    
    HResDisp = DispCallFunc(pUnk, vtblIdx * 4, 4, vbLong, i, VType(0), VPtr(0), vtblCall)
    If HResDisp <> S_OK Then Err.Raise HResDisp, , "Error in DispCallFunc"
End Function
