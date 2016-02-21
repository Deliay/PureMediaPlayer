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

Private m_CurSelFilter As IFilterInfo, m_CurSelPin As IPinInfo
Private LAVVideoIndex As Long, LAVAudioIndex As Long
Private LAVSplitterIndex As Long, LAVSplitterSourceIndex As Long
Private VSFilterIndex As Long

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
    Dim m_GraphManager As New FilgraphManager
    FillDecoder m_GraphManager
    
    If (LAVAudioIndex = -1 Or _
        LAVVideoIndex = -1 Or _
        LAVSplitterIndex = -1 Or _
        VSFilterIndex = -1 Or _
        LAVSplitterSourceIndex = -1) Then
        
        RegisterAllDecoder
        Set m_GraphManager = New m_GraphManager
        FillDecoder m_GraphManager
    End If
    
    If (LAVAudioIndex = -1 Or _
        LAVVideoIndex = -1 Or _
        LAVSplitterIndex = -1 Or _
        VSFilterIndex = -1 Or _
        LAVSplitterSourceIndex = -1) Then
    
        MsgBox "Cannot Register Decoder! Please run me with higher"
    End If
    
    BuildGrph "E:\@ ”∆µ\@”Œœ∑\Films\rec.mp4", m_GraphManager
    
End Sub

Public Function BuildGrph(ByVal srcFile As String, ByVal m_GraphManager As FilgraphManager)
    Dim objSrcSplitterReg As IRegFilterInfo, objSrcSplitterFilter As IFilterInfo
    Dim objVideoReg As IRegFilterInfo, objVideoFilter As IFilterInfo
    Dim objAudioReg As IRegFilterInfo, objAudioFilter As IFilterInfo
    Dim FileName As String
    
    m_GraphManager.RegFilterCollection.Item LAVSplitterSourceIndex, objSrcSplitterReg
    objSrcSplitterReg.Filter objSrcSplitterFilter
    
    CheckForFileSinkAndSetFileName objSrcSplitterFilter, srcFile
    
    For Each objPinInfo In objSrcSplitterFilter.Pins
        
    Next
End Function



Private Sub FillDecoder(m_GraphManager As FilgraphManager)
    Dim i As Long
    Dim objRegFilter As Object
    For Each objRegFilter In m_GraphManager.RegFilterCollection
        i = i + 1
        Debug.Print objRegFilter.Name
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

Public Function CheckForFileSinkAndSetFileName(ByVal Flt As olelib.IUnknown, FileName As String) As Boolean
    Const IID_IFileSinkFilter = "{A2104830-7C70-11CF-8BCE-00AA00A3F1A6}", VTbl_SetFileName = 3
    Dim oUnkFSink As stdole.IUnknown

    Set oUnkFSink = CastToUnkByIID(Flt, IID_IFileSinkFilter)
    CheckForFileSinkAndSetFileName = vtblCall(ObjPtr(oUnkFSink), VTbl_SetFileName, StrPtr(FileName), 0&) = S_OK
End Function

Public Function CastToUnkByIID(ByVal ObjToCastFrom As olelib.IUnknown, IID As String) As stdole.IUnknown
    Dim UUID As olelib.UUID
    olelib.CLSIDFromString IID, UUID
    ObjToCastFrom.QueryInterface UUID, CastToUnkByIID
End Function

Public Function vtblCall(ByVal pUnk As Long, ByVal vtblIdx As Long, ParamArray P() As Variant)
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
