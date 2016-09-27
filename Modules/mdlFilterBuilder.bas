Attribute VB_Name = "mdlFilterBuilder"
Option Explicit

Global IsIDE     As Boolean

Global IsRestart As Boolean

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

Private Declare Function csri_renderer_default Lib "vsfilter.dll" () As IntPtr

Private Declare Function csri_open_file _
                Lib "vsfilter.dll" (Renderer As IntPtr, _
                                    FileName As IntPtr, _
                                    flags As csri_openflag) As IntPtr

Private Type csri_openflag

    Name As IntPtr
    
End Type

Private objSrcSplitterReg As IRegFilterInfo, objSrcSplitterFilter As IFilterInfo

Private objVideoReg       As IRegFilterInfo, objVideoFilter As IFilterInfo, objVideoPin As IPinInfo

Private objAudioReg       As IRegFilterInfo, objAudioFilter As IFilterInfo, objAudioPin As IPinInfo

Private objSubtitleReg    As IRegFilterInfo, objSubtitleFilter As IFilterInfo, objSubtitlePin As IPinInfo

Private objRenderReg      As IRegFilterInfo, objRenderFilter As IFilterInfo ', objRenderPin As IPinInfo

Private objSpliterPin     As IPinInfo

Private LAVVideoIndex     As Long, LAVAudioIndex As Long

Private LAVSplitterIndex  As Long, LAVSplitterSourceIndex As Long

Private VSFilterIndex     As Long, EVRIndex As Long, MadVRIndex As Long

Private VMR9Index         As Long, VMR7Index As Long, VRIndex As Long

Public EVRFilterStorage   As IBaseFilter

Private pAdminGroup       As IntPtr

Type SID_IDENTIFIER_AUTHORITY

    Value(6) As Byte

End Type

Private Const SECURITY_BUILTIN_DOMAIN_RID As Long = &H20

Private Const DOMAIN_ALIAS_RID_ADMINS     As Long = &H220

Private Const NULL_                       As Long = 0

Private Const SECURITY_NT_AUTHORITY       As Long = &H5

Private Declare Function AllocateAndInitializeSid _
                Lib "Advapi32.dll" (SID As SID_IDENTIFIER_AUTHORITY, _
                                    ByVal Count As Byte, _
                                    ByVal dwSubAuth0 As Long, _
                                    ByVal dwSubAuth1 As Long, _
                                    ByVal dwSubAuth2 As Long, _
                                    ByVal dwSubAuth3 As Long, _
                                    ByVal dwSubAuth4 As Long, _
                                    ByVal dwSubAuth5 As Long, _
                                    ByVal dwSubAuth6 As Long, _
                                    ByVal dwSubAuth7 As Long, _
                                    vpPSID As IntPtr) As BOOL
                                    
Private Declare Function CheckTokenMembership _
                Lib "Advapi32.dll" (ByVal hToken As IntPtr, _
                                    ByVal vpPSID As IntPtr, _
                                    ByRef isMember As BOOL) As BOOL

Private Declare Function FreeSid Lib "Advapi32.dll" (vpPSID As Long) As Long
                
Private Declare Function ShellExecuteEx _
                Lib "shell32" _
                Alias "ShellExecuteExW" (SEI As SHELLEXECUTEINFO2) As Long

Private Declare Function WaitForSingleObject _
                Lib "kernel32" (ByVal hHandle As Long, _
                                ByVal dwMilliseconds As Long) As Long

Private Const INFINITE = &HFFFF      '  Infinite timeout

Private Declare Function SetProcessDPIAware Lib "user32" () As BOOL

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Public isAdminPerm As Boolean

Public Sub RegisterCOM()
    ReqAdminPerm
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\SSubTmr6.dll", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\vbaListView6.ocx", App.Path & "\", 0
    RegisterSSubTmr6
    RegistervbaListView6
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\SSubTmr6.dll", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\vbaListView6.ocx", App.Path & "\", 0
    
End Sub

Public Sub RegisterAllDecoder()
    ReqAdminPerm
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
    IsIDE = GetIDEmode
    InitPerm
    
    InitConfigFiles

    'SetProcessDpiAwareness_soft
    If (IsSupportDIPSet) Then
        SetProcessDPIAware
        SetProcessDpiAwareness_hard

    End If

    If (Len(Command) = 6 And Left$(Command, 6) = "--perm") Then
        RegisterAllDecoder
        RegisterCOM
        End

    End If

    If (Len(Command) = 9 And Left$(Command, 9) = "--restart") Then
        IsRestart = True
        GoTo FillDecoder

    End If
    
    If (App.PrevInstance) Then
        If (Len(Command) <> 0) Then

            Dim cbs    As COPYDATASTRUCT

            Dim ptrCmd As IntPtr

            ptrCmd = GetCommandLine()
            cbs.dwData = 3
            cbs.cbData = (lstrlen(ptrCmd) + 1) * 2
            cbs.lpData = ptrCmd
            SendMessageW val(GlobalConfig.LastHwnd), WM_COPYDATA, ByVal 0&, cbs

        End If
        
        SendMessageW val(GlobalConfig.LastHwnd), PM_ACTIVE, 0&, 0&
        End
        Exit Sub

    End If

FillDecoder:
    LAVVideoIndex = -1
    LAVAudioIndex = -1
    LAVSplitterIndex = -1
    LAVSplitterSourceIndex = -1
    VSFilterIndex = -1
    VMR9Index = -1
    EVRIndex = -1
    MadVRIndex = -1

    Dim lngRertyCount As Long

ReFill:
    FillDecoder mdlGlobalPlayer.GlobalFilGraph

    If (LAVAudioIndex = -1 Or LAVVideoIndex = -1 Or LAVSplitterIndex = -1 Or VSFilterIndex = -1 Or LAVSplitterSourceIndex = -1 Or MadVRIndex = -1) Then
        ReqAdminPerm
        Shell App.Path & "\" & App.EXEName & ".exe --restart", vbNormalFocus
        End
        GoTo ReFill

    End If
    
    If (VMR9Index = -1 And EVRIndex = -1 And VMR7Index = -1) Then
        MsgBox "Bad Renderer Support!"
        End

    End If
    
    If (LAVAudioIndex = -1 Or LAVVideoIndex = -1 Or LAVSplitterIndex = -1 Or VSFilterIndex = -1 Or LAVSplitterSourceIndex = -1 Or MadVRIndex = -1) Then
    
        MsgBox "Can't register decodes! Please allow permission request in UAC or other Security Software"
        End

    End If

    Load frmMenu

    If (VMR9Index = -1) Then
        frmMenu.Renderers(RenderType.VideoMixedRenderer9).Enabled = False

    End If

    If (EVRIndex = -1) Then
        frmMenu.Renderers(RenderType.EnhancedVideoRenderer).Enabled = False

    End If

    If (GlobalConfig.Renderer = "") Then
        frmMenu.Renderers_Click RenderType.MadVRednerer
    Else
        frmMenu.Renderers_Click val(GlobalConfig.Renderer)

    End If

    'Create A Clean FilgraphManager
    Set mdlGlobalPlayer.GlobalFilGraph = New FilgraphManager
    
    'But Issue:
    'Don't Register COM DLL whitout error
    'If (Not IsIDE) Then RegisterCOM
    
    On Error GoTo RegisterCOMErr

    mdlToolBarAlphaer.LoadUI
    
    If (Not IsIDE) Then StartHook frmMain.hWnd
    
    If (isAdminPerm) Then
        frmMain.Caption = "¹ÜÀíÔ±: " & frmMain.Caption

    End If

    If (IsIDE) Then
        frmMain.Caption = "(IDE) " & frmMain.Caption

    End If

    On Error GoTo 0

    Dim i As Variant

    For Each i In GlobalConfig.LastPlayList

        If (i = "@") Then GoTo Placement
        mdlPlaylist.AddFileToPlaylist i
Placement:
    Next

    If (Not IsRestart And Not IsIDE) Then InitialCommandLine
    
    Exit Sub
RegisterCOMErr:
    'do gui com register
    RegisterCOM
    Shell App.Path & "\" & App.EXEName & ".exe --restart", vbNormalFocus
    End
    Resume

End Sub

Public Sub InitPerm()

    Dim isAdminPermission As BOOL
    
    Dim NtAuthority       As SID_IDENTIFIER_AUTHORITY
    
    NtAuthority.Value(5) = SECURITY_NT_AUTHORITY
    isAdminPermission = AllocateAndInitializeSid(NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, pAdminGroup)
    
    If (isAdminPermission) Then
        If (Not CheckTokenMembership(NULL_, pAdminGroup, isAdminPermission) = 1) Then
            CheckTokenMembership NULL_, pAdminGroup, isAdminPermission
            Debug.Print GetLastError
            isAdminPermission = False

        End If

        FreeSid pAdminGroup

    End If

    isAdminPerm = Not (isAdminPermission = 0)

End Sub

Public Function ReqAdminPerm()
    
    If (isAdminPerm = False) Then

        Dim sLInfo As SHELLEXECUTEINFO2

        With sLInfo
            .cbSize = Len(sLInfo)
            .lpVerb = StrPtr("runas")
            .lpFile = StrPtr(App.Path & "\" & App.EXEName & ".exe")
            .hWnd = 0
            .nShow = 1
            .lpParameters = StrPtr("--perm")
            
        End With
        
        If (ShellExecuteEx(sLInfo) = 0) Then
            MsgBox "You deny the permission request! Exit."
            End

        End If
        
        WaitForSingleObject sLInfo.hProcess, INFINITE
        Sleep 1000

    End If
    
End Function

Public Sub BuildGrph(ByVal srcFile As String, _
                     ByRef objGraphManager As FilgraphManager, _
                     ByRef boolHasVideo As Boolean, _
                     ByRef boolHasAudio As Boolean, _
                     ByRef boolHasSubtitle As Boolean, _
                     Optional ByVal eRenderer As RenderType = VideoMixedRenderer9)
    boolHasSubtitle = False
    'First put subtitle flag to FALSE

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
    '3. LAVSplit (Audio) -> (Input) LAVAudio (Output) -> (Video) VSFilter (Output)-> Connect Downstream
    
    '  Audio first
    If (boolHasAudio = True) Then objAudioPin.Render 'audio auto connect
    
    'If Splitter not give a subtitle pin, Create a vs-fliter and connect
    If ((boolHasVideo = True) And (boolHasSubtitle = False)) Then
        objGraphManager.RegFilterCollection.Item VSFilterIndex, objSubtitleReg
        objSubtitleReg.Filter objSubtitleFilter

    End If
    
    '  Video second
    If (boolHasVideo = True) Then

        ' LAVSplit (Subtitle) -> (Input) VSFilter
        If (boolHasSubtitle = True) Then

            Dim objPinVSInput As IPinInfo

            For Each objPinVSInput In objSubtitleFilter.Pins

                If (LCase(objPinVSInput.Name) = "input") Then Exit For
            Next
            objSubtitlePin.Connect objPinVSInput
            objVideoPin.Render
        Else

            'Force Connect Here
            Dim objPinForceVSInput As IPinInfo

            For Each objPinForceVSInput In objSubtitleFilter.Pins

                If (LCase$(objPinForceVSInput.Name) = "video") Then Exit For
            Next
            objVideoPin.Connect objPinForceVSInput

            Dim objRendererInput As IPinInfo

            For Each objRendererInput In objRenderFilter.Pins

                If (InStr(1, LCase$(objRendererInput.Name), "input") <> 0) Then Exit For
            Next

            Dim objVSFilterOutput As IPinInfo

            For Each objVSFilterOutput In objSubtitleFilter.Pins

                If (LCase$(objVSFilterOutput.Name) = "output") Then Exit For
            Next
            objVSFilterOutput.Connect objRendererInput

        End If

        ' Connect it to Splitter Subtitle
        ' In the end ,render
        
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

Private Function CastToIUnknow(ByVal Flt As olelib.IUnknown) As olelib.IUnknown
    Set CastToIUnknow = Flt

End Function

Public Function SetVSFilterFileName(FileName As String) As Boolean

    Const IID_IDirectVobSub = "{EBE1FB08-3957-47ca-AF13-5827E5442E56}", VTbl_SetFileName = 4

    Dim oDirectVobSub As stdole.IUnknown

    Set oDirectVobSub = CastToUnkByIID(objSubtitleFilter.Filter, IID_IDirectVobSub)
    SetVSFilterFileName = vtblCall(ObjPtr(oDirectVobSub), VTbl_SetFileName, vbEmpty, StrPtr(FileName)) = S_OK
    
    If (GlobalConfig.SubtitleBind.Exist(mdlGlobalPlayer.FileMD5)) Then
        GlobalConfig.SubtitleBind.Value(mdlGlobalPlayer.FileMD5) = FileName
    Else
        GlobalConfig.SubtitleBind.AddKeyValue mdlGlobalPlayer.FileMD5, FileName

    End If

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

    If (IsEmpty(objVideoFilter) Or objVideoFilter Is Nothing) Then
        mdlGlobalPlayer.GlobalFilGraph.RegFilterCollection.Item LAVVideoIndex, objVideoReg
        objVideoReg.Filter objVideoFilter

    End If

    ShowPropertyPage objVideoFilter.Filter, "PureMediaPlayer - " & objVideoFilter.Name, frmMain.hWnd

End Function

Public Function ShowAudioDecoderConfig()

    If (IsEmpty(objAudioFilter) Or objAudioFilter Is Nothing) Then
        mdlGlobalPlayer.GlobalFilGraph.RegFilterCollection.Item LAVAudioIndex, objAudioReg
        objAudioReg.Filter objAudioFilter

    End If

    ShowPropertyPage objAudioFilter.Filter, "PureMediaPlayer - " & objAudioFilter.Name, frmMain.hWnd

End Function

Public Function ShowSpliterConfig()

    If (IsEmpty(objSrcSplitterFilter) Or objSrcSplitterFilter Is Nothing) Then
        mdlGlobalPlayer.GlobalFilGraph.RegFilterCollection.Item LAVSplitterIndex, objSrcSplitterReg
        objSrcSplitterReg.Filter objSrcSplitterFilter

    End If

    ShowPropertyPage objSrcSplitterFilter.Filter, "PureMediaPlayer - " & objSrcSplitterFilter.Name, frmMain.hWnd

End Function

Public Function ShowSubtitleConfig()

    If (IsEmpty(objSubtitleFilter) Or objSubtitleFilter Is Nothing) Then

        mdlGlobalPlayer.GlobalFilGraph.RegFilterCollection.Item VSFilterIndex, objSubtitleReg
        objSubtitleReg.Filter objSubtitleFilter

    End If

    ShowPropertyPage objSubtitleFilter.Filter, "PureMediaPlayer - " & objSubtitleFilter.Name, frmMain.hWnd

End Function

Public Function ShowRendererConfig()

    If (IsEmpty(objRenderFilter) Or objRenderFilter Is Nothing) Then
    
        Dim VMRVerSpec As Long
    
        If (mdlConfig.GlobalConfig.Renderer = VideoMixedRenderer9) Then
            VMRVerSpec = VMR9Index
            
        ElseIf (mdlConfig.GlobalConfig.Renderer = VideoMixedRenderer) Then
            VMRVerSpec = VMR7Index
            
        ElseIf (mdlConfig.GlobalConfig.Renderer = EnhancedVideoRenderer) Then
            VMRVerSpec = EVRIndex
            
        ElseIf (mdlConfig.GlobalConfig.Renderer = MadVRednerer) Then
            VMRVerSpec = MadVRIndex
            
        Else
            VMRVerSpec = VRIndex
    
        End If
        
        mdlGlobalPlayer.GlobalFilGraph.RegFilterCollection.Item VMRVerSpec, objRenderReg
        objRenderReg.Filter objRenderFilter

    End If
    
    ShowPropertyPage objRenderFilter.Filter, "PureMediaPlayer - " & objRenderFilter.Name, frmMain.hWnd

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
