Attribute VB_Name = "mdlGlobalPlayer"
Option Explicit

Public Enum RenderType

    VideoRenderer
    VideoMixedRenderer
    VideoMixedRenderer9
    EnhancedVideoRenderer
    MadVRednerer

End Enum

Public Const DefaultTitle As String = "Pure Media Player"

Public GlobalFilGraph     As New FilgraphManager

Private lngStorageH       As Long, lngStorageW As Long

Private lngStorageT       As Long, lngStorageL As Long

Private ifPostion         As IMediaPosition

Private ifPlayback        As IVideoWindow

Private ifVideo           As IBasicVideo

Private ifVolume          As IBasicAudio

Private ifType            As IMediaTypeInfo

Public GlobalRenderType   As RenderType

Public EVRHoster          As EVRRenderer

Public Width              As Long

Public Height             As Long

Private boolLoadedFile    As Boolean

Private boolIsNetFile     As Boolean

Private boolIsFullScreen  As Boolean

Private strLastestFile    As String

Private VideoRatio        As Double

Private hasVideo_         As Boolean

Private hasAudio_         As Boolean

Private hasSubtitle_      As Boolean

Private strFileOpenedMD5  As String

Public Enum StatusBarEnum

    Action = 1
    PlayBack = 2
    PlayTime = 3
    FileName = 4
    
End Enum

Public Enum PlayStatus

    Stopped
    Paused
    Running
    Caching

End Enum

Private Const WS_BORDER = &H800000

Private Const WS_CAPTION = &HC00000

Private Const WS_THICKFRAME = &H40000

Private Const WS_SIZEBOX = WS_THICKFRAME

Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Public Property Get IsNetFile() As Boolean
    IsNetFile = boolIsNetFile

End Property

Public Property Get IsWindowed() As Boolean
    IsWindowed = Not boolIsFullScreen

End Property

Public Property Get FileMD5() As String
    FileMD5 = strFileOpenedMD5

End Property

Public Property Get GlobalPlayStatus() As PlayStatus

    GlobalFilGraph.GetState 500, GlobalPlayStatus

End Property

Public Property Let Volume(v As Long)

    If (v > 100) Or (v < 0) Then Exit Property
    ifVolume.Volume = -((100 - v) * 100)

End Property

Public Property Get Volume() As Long

    If (ifVolume Is Nothing And Not mdlGlobalPlayer.hasAudio) Then Exit Property
    Volume = 100 + ifVolume.Volume / 100
    UpdateStatus StaticString(PLAYER_VIOCE_RATE) & ":" & Volume, PlayBack

End Property

Public Property Get Precent() As Single
    
    If ifPostion Is Nothing Then Exit Property
    Precent = (CurrentTime / Duration) * 100
    
End Property

Public Property Get Rate() As Long
    '0.5 per level
    'rate 100 mean 1
    
    Rate = ifPostion.Rate * 100
    
End Property

Public Property Let Rate(v As Long)

    '0.5 per level
    'rate 100 mean 1
    If (v > 400) Then Exit Property
    ifPostion.Rate = v / 100

End Property

Public Property Let Precent(Value As Single)
    CurrentTime = Duration * (Value / 100)
    
End Property

Public Property Get Loaded() As Boolean
    Loaded = boolLoadedFile

End Property

Public Property Get File() As String
    
    File = strLastestFile

End Property

Public Property Let File(v As String)
    SaveCurrentPos
    strLastestFile = v
    strFileOpenedMD5 = MD5String(v)

End Property

Public Property Get hasVideo() As Boolean
    
    hasVideo = hasVideo_
    
End Property

Public Property Get hasAudio() As Boolean

    hasAudio = hasAudio_

End Property

Public Property Get hasSubtitle() As Boolean

    hasSubtitle = hasSubtitle_

End Property

Public Sub RenderMediaFile()

    Dim strFilePath As String

    Set GlobalFilGraph = Nothing
    Set GlobalFilGraph = New FilgraphManager

    UpdateStatus StaticString(PLAYER_STATUS_LOADING), Action

    strFilePath = File
    UpdateStatus NameGet(strFilePath), FileName
    
    hasVideo_ = False: hasAudio_ = False: hasSubtitle_ = False

    GlobalRenderType = val(GlobalConfig.Renderer)

    If (Not GlobalFilGraph Is Nothing) Then
        mdlGlobalPlayer.CloseFile
        Set GlobalFilGraph = Nothing
    End If
    'mdlFilterBuilder.BuildGrph strFilePath, GlobalFilGraph, hasVideo_, hasAudio_, hasSubtitle_, GlobalRenderType
    mdlFilterProductor.BuildGraph strFilePath, GlobalFilGraph, hasVideo_, hasAudio_, hasSubtitle_, GlobalRenderType
    boolLoadedFile = True

    If (hasVideo = False And hasAudio_ = False) Then GoTo DcodeErr
    UpdateTitle NameGet(strFilePath)
    Set ifPostion = GlobalFilGraph

    If (GlobalRenderType = EnhancedVideoRenderer) Then
        Set EVRHoster = New EVRRenderer
        EVRHoster.CreateInterface mdlFilterProductor.EVRFilterStorage

    End If

    If (hasVideo) Then
        If (GlobalRenderType <> EnhancedVideoRenderer) Then
            Set ifVideo = GlobalFilGraph
            Set ifPlayback = GlobalFilGraph

            'ifPlayback.Caption = "PureMediaPlayer - LayerWindow"
            ifPlayback.Owner = frmMain.frmPlayer.hWnd
            ifPlayback.MessageDrain = frmMain.frmPlayer.hWnd
            
            Dim lngSrcStyle As Long
            
            lngSrcStyle = ifPlayback.WindowStyle
            lngSrcStyle = lngSrcStyle And Not WS_BORDER
            lngSrcStyle = lngSrcStyle And Not WS_CAPTION
            lngSrcStyle = lngSrcStyle And Not WS_SIZEBOX

            ifPlayback.WindowStyle = lngSrcStyle
            mdlGlobalPlayer.ResizePlayWindow

            ifPlayback.HideCursor False
        Else
            EVRHoster.SetPlayBackWindow frmMain.frmPlayer.hWnd
            mdlGlobalPlayer.ResizePlayWindow

        End If
        
    End If
    
    If (hasAudio) Then
        Set ifVolume = GlobalFilGraph

    End If
 
    mdlGlobalPlayer.CurrentTime = 0
    mdlPlaylist.SetItemLength strFilePath, FormatedDuration
    
    If Not frmMain.NowPlaying Is Nothing Then frmMain.NowPlaying.NowPlaying = False
    
    On Error GoTo notLoadPlaylist
    
    Set frmMain.NowPlaying = frmMain.lstPlaylist(strFilePath)
    
notLoadPlaylist:

    Resume Next

    If Not frmMain.NowPlaying Is Nothing Then frmMain.NowPlaying.NowPlaying = True
    
    boolLoadedFile = True
    
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    
hErr:

    Resume Next

    mdlGlobalPlayer.CurrentTime = 0
    SeekCurrentPos
    mdlGlobalPlayer.Play
    PlayPauseSwitch

    DoEvents
    mdlGlobalPlayer.ResizePlayWindow

    Dim strFileMD5 As String

    If (GlobalConfig.SubtitleBind.Exist(mdlGlobalPlayer.FileMD5)) Then
        mdlFilterProductor.SetVSFilterFileName GlobalConfig.SubtitleBind(mdlGlobalPlayer.FileMD5)

    End If

    mdlToolBarAlphaer.SwitchUI True, False
    mdlToolBarAlphaer.SwitchUI True, True
    If (mdlToolBarAlphaer.boolPlaylistStatus) Then mdlToolBarAlphaer.PlaylistShow
    
    QueryMediaStreams
    
    Exit Sub

DcodeErr:
    MsgBox mdlLanguageApplyer.StaticString(TIPS_NOT_SUPPORT)
    mdlPlaylist.SetItemLength File, mdlLanguageApplyer.StaticString(FILE_NOT_SUPPORT)

    Exit Sub

    MsgBox mdlLanguageApplyer.StaticString(TIPS_UNKNOW_ERR)

    Exit Sub
    
End Sub

Public Property Get CurrentTime() As Double
    
    If ifPostion Is Nothing Then Exit Property
    CurrentTime = ifPostion.CurrentPosition
    
End Property

Public Property Get FormatedCurrentTime() As String

    If ifPostion Is Nothing Then Exit Property

    FormatedCurrentTime = (CurrentTime \ 60) & ":" & (CurrentTime Mod 60)

End Property

Public Property Let CurrentTime(val As Double)
    
    If ifPostion Is Nothing Then Exit Property
    ifPostion.CurrentPosition = val
    
End Property

Public Property Get Duration() As Double

    If ifPostion Is Nothing Then Exit Property
    Duration = ifPostion.Duration

End Property

Public Property Get FormatedDuration() As String

    FormatedDuration = (Duration \ 60) & ":" & (Duration Mod 60)

End Property
'
'
'Public Sub RaiseRegFileter(List As vbalListViewCtl)
'
'    Dim objRegFilter As IRegFilterInfo
'
'    Dim objFilter    As IFilterInfo
'
'    If (GlobalFilGraph Is Nothing) Then Set GlobalFilGraph = New FilgraphManager
'
'    For Each objRegFilter In GlobalFilGraph.RegFilterCollection
'
'        List.ListItems.Add , , objRegFilter.Name
'    Next
'
'End Sub

Public Sub Play()
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    UpdateStatus StaticString(PLAY_STATUS_PLAYING), PlayBack
    
hErr:

    GlobalFilGraph.Run
    
    frmMain.tmrUpdateTime.Enabled = True
    
    If (Duration = 0) Then frmMain.tmrUpdateTime.Enabled = False
    
    SaveCurrentPos

End Sub

Public Sub Pause()
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub

    GlobalFilGraph.Pause

    UpdateStatus StaticString(PLAY_STATUS_PAUSED), PlayBack
    
    SaveCurrentPos

End Sub

Public Sub StopPlay()
    SaveCurrentPos
    Precent = 0
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    
    If Not GlobalFilGraph Is Nothing And Len(File) <> 0 Then

        GlobalFilGraph.Stop

    End If
    
End Sub

Public Sub SwitchPlayStauts()
    
    If (mdlGlobalPlayer.Loaded) Then
        If (GlobalPlayStatus = Running) Then
            mdlGlobalPlayer.Pause
        ElseIf (GlobalPlayStatus = Paused) Then
            mdlGlobalPlayer.Play
            
        End If
        
    End If

    PlayPauseSwitch

End Sub

Public Sub CloseFile()
    SaveCurrentPos
    UpdateStatus StaticString(PLAYER_STATUS_CLOSEING), Action
    strLastestFile = ""
    UpdateTitle StaticString(PLAYER_STATUS_IDLE)
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    mdlGlobalPlayer.Pause
    mdlGlobalPlayer.StopPlay
    
End Sub

Public Sub ResizePlayWindow()
    
    If (Not hasVideo_) Then Exit Sub
    
    If (ifVideo Is Nothing) Then
        If (GlobalRenderType <> EnhancedVideoRenderer) Then

            Exit Sub

        End If

    End If

    'On Error GoTo hErr
    
    Dim commonW As Long, commonH As Long
    
    Dim resultW As Long, resultH As Long
    
    Dim resultT As Long, resultL As Long

    If (GlobalRenderType = EnhancedVideoRenderer) Then
        EVRHoster.GetVideoSize commonW, commonH
    Else

        ifVideo.GetVideoSize commonW, commonH

    End If

    VideoRatio = commonW / commonH
    resultW = Width
    resultH = Width / VideoRatio
    
    If (resultH > Height) Then
        resultW = VideoRatio * Height
        resultH = Height
        
    End If
    
    resultT = (Height - resultH) / 2
    resultL = (Width - resultW) / 2

    If (GlobalRenderType = EnhancedVideoRenderer) Then
        EVRHoster.SetVideoSize resultL, resultT, resultW, resultH
    Else

        ifPlayback.SetWindowPosition resultL, resultT, resultW, resultH
        
    End If

hErr:
    frmMain.pbTimeBar.ZOrder 0
    
End Sub

Public Sub UpdateStatus(strCaption As String, Target As StatusBarEnum)

    If (Target = PlayBack) Then
        frmMain.Label2.Caption = strCaption
    ElseIf Target = PlayTime Then
        frmMain.Label1.Caption = strCaption

    End If
    
End Sub

Public Sub UpdateTitle(strCaption As String)
    frmMain.Caption = DefaultTitle & " - " & strCaption

End Sub

Public Sub RaiseMediaFilter(List As ListBox)
    List.AddItem mdlGlobalPlayer.File

    Dim objFilter As IFilterInfo
    
    If (GlobalFilGraph Is Nothing) Then Exit Sub
    If (GlobalFilGraph.FilterCollection Is Nothing) Then Exit Sub
    
    For Each objFilter In GlobalFilGraph.FilterCollection
        
        List.AddItem objFilter.Name
        
    Next
    
End Sub

Public Sub SwitchFullScreen(Optional force As Boolean = False, _
                            Optional forceValue As Boolean = False)

    If (Not hasVideo) Then Exit Sub
    If (ifVideo Is Nothing) Then
        If (GlobalRenderType <> EnhancedVideoRenderer) Then

            Exit Sub

        End If

    End If

    If force = True Then
        boolIsFullScreen = forceValue
        ResizeFullScreen

        Exit Sub

    End If
    
    boolIsFullScreen = Not boolIsFullScreen
    ResizeFullScreen

End Sub

Public Sub ResizeFullScreen()

    If (boolIsFullScreen) Then
        lngStorageW = frmMain.Width
        lngStorageH = frmMain.Height
        lngStorageT = frmMain.Top
        lngStorageL = frmMain.Left
        frmMain.BorderStyle = 0
        UpdateTitle NameGet(File)
        frmMain.WindowState = 2
    Else
    
        frmMain.BorderStyle = 2
        UpdateTitle NameGet(File)
        frmMain.WindowState = 0

        If (lngStorageW <> 0) Then
            frmMain.Width = lngStorageW
            frmMain.Height = lngStorageH
            frmMain.Top = lngStorageT
            frmMain.Left = lngStorageL

        End If

    End If
    
End Sub

Public Sub SeekLastPos(ByVal strMD5 As String)
    mdlGlobalPlayer.CurrentTime = val(GlobalConfig.LastPlayPos(strMD5))
    
    If (mdlGlobalPlayer.CurrentTime = mdlGlobalPlayer.Duration) Then
        mdlGlobalPlayer.CurrentTime = 0
        
    End If
    
End Sub

Public Sub SeekCurrentPos()
    SeekLastPos mdlGlobalPlayer.FileMD5

End Sub

Public Sub SaveCurrentPos()

    If (mdlGlobalPlayer.Loaded) Then

        GlobalConfig.LastPlayPos.Value(mdlGlobalPlayer.FileMD5) = CStr(mdlGlobalPlayer.CurrentTime)

    End If

End Sub
