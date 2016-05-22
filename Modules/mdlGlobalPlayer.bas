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

Private EVRHoster         As EVRRenderer

Public Width              As Long

Public Height             As Long

Private boolLoadedFile    As Boolean

Private boolIsFullScreen  As Boolean

Private strLastestFile    As String

Private VideoRatio        As Double

Private hasVideo_         As Boolean

Private hasAudio_         As Boolean

Private hasSubtitle_      As Boolean

Public Enum StatusBarEnum

    Action = 1
    PlayBack = 2
    PlayTime = 3
    FileName = 4
    
End Enum

Public Enum PlayStatus

    playing
    Paused
    Stoped

End Enum

Public GlobalPlayStatus As PlayStatus

Private Const WS_BORDER = &H800000

Private Const WS_CAPTION = &HC00000

Private Const WS_THICKFRAME = &H40000

Private Const WS_SIZEBOX = WS_THICKFRAME

Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)

Public Property Get IsWindowed() As Boolean
    IsWindowed = Not boolIsFullScreen

End Property

Public Property Let Volume(V As Long)

    If (V > 100) Or (V < 0) Then Exit Property
    ifVolume.Volume = -((100 - V) * 100)

End Property

Public Property Get Volume() As Long

    If (ifVolume Is Nothing And Not mdlGlobalPlayer.HasAudio) Then Exit Property
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

Public Property Let Rate(V As Long)

    '0.5 per level
    'rate 100 mean 1
    If (V > 400) Then Exit Property
    ifPostion.Rate = V / 100

End Property

Public Property Let Precent(value As Single)
    CurrentTime = Duration * (value / 100)
    
End Property

Public Property Get Loaded() As Boolean
    Loaded = boolLoadedFile

End Property

Public Property Get File() As String
    
    File = strLastestFile

End Property

Public Property Let File(V As String)
    SaveCurrentPos
    strLastestFile = V

End Property

Public Property Get HasVideo() As Boolean
    
    HasVideo = hasVideo_
    
End Property

Public Property Get HasAudio() As Boolean

    HasAudio = hasAudio_

End Property

Public Property Get HasSubtitle() As Boolean

    HasSubtitle = hasSubtitle_

End Property

Public Sub RenderMediaFile()

    Dim strFilePath As String

    Set GlobalFilGraph = Nothing
    Set GlobalFilGraph = New FilgraphManager

    UpdateStatus StaticString(PLAYER_STATUS_LOADING), Action

    strFilePath = File
    UpdateStatus Dir(strFilePath), FileName
    
    hasVideo_ = False: hasAudio_ = False: hasSubtitle_ = False
    GlobalRenderType = val(getConfig(CFG_SETTING_RENDERER))
    mdlFilterBuilder.BuildGrph strFilePath, GlobalFilGraph, hasVideo_, hasAudio_, hasSubtitle_, GlobalRenderType
    
    If (HasVideo = False And hasAudio_ = False) Then GoTo DcodeErr
    UpdateTitle Dir(File)
    Set ifPostion = GlobalFilGraph

    If (GlobalRenderType = EnhancedVideoRenderer) Then
        Set EVRHoster = New EVRRenderer
        EVRHoster.CreateInterface mdlFilterBuilder.EVRFilterStorage

    End If

    If (HasVideo) Then
        If (GlobalRenderType <> EnhancedVideoRenderer) Then
            Set ifVideo = GlobalFilGraph
            Set ifPlayback = GlobalFilGraph
            'ifPlayback.Caption = "PureMediaPlayer - LayerWindow"
            ifPlayback.Owner = frmMain.frmPlayer.hwnd
            ifPlayback.MessageDrain = frmMain.frmPlayer.hwnd
            
            Dim lngSrcStyle As Long
            
            lngSrcStyle = ifPlayback.WindowStyle
            lngSrcStyle = lngSrcStyle And Not WS_BORDER
            lngSrcStyle = lngSrcStyle And Not WS_CAPTION
            lngSrcStyle = lngSrcStyle And Not WS_SIZEBOX
            ifPlayback.WindowStyle = lngSrcStyle
            mdlGlobalPlayer.ResizePlayWindow
            ifPlayback.HideCursor False
        Else
            EVRHoster.SetPlayBackWindow frmMain.frmPlayer.hwnd
            mdlGlobalPlayer.ResizePlayWindow

        End If
        
    End If
    
    If (HasAudio) Then
        Set ifVolume = GlobalFilGraph

    End If
 
    mdlGlobalPlayer.CurrentTime = 0
    mdlPlaylist.SetItemLength strFilePath, FormatedDuration
    
    If Not frmMain.nowPlaying Is Nothing Then frmMain.nowPlaying.ForeColor = vbWhite
    
    On Error GoTo notLoadPlaylist
    
    Set frmMain.nowPlaying = frmMain.lstPlaylist.ListItems(strFilePath)
    
notLoadPlaylist:

    Resume Next

    If Not frmMain.nowPlaying Is Nothing Then frmMain.nowPlaying.ForeColor = vbGrayText
    
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
    
    Exit Sub
DcodeErr:
    MsgBox "Not Support this codes type yet!"
    mdlPlaylist.SetItemLength File, mdlLanguageApplyer.StaticString(FILE_NOT_SUPPORT)
    Exit Sub

    MsgBox "Unknow Err"
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

Public Sub RaiseRegFileter(list As vbalListViewCtl)
    
    Dim objRegFilter As IRegFilterInfo
    
    Dim objFilter    As IFilterInfo
    
    If (GlobalFilGraph Is Nothing) Then Set GlobalFilGraph = New FilgraphManager
    
    For Each objRegFilter In GlobalFilGraph.RegFilterCollection
        
        list.ListItems.Add , , objRegFilter.Name
    Next
    
End Sub

Public Sub Play()
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    UpdateStatus StaticString(PLAY_STATUS_PLAYING), PlayBack
    
hErr:
    GlobalFilGraph.Run
    GlobalPlayStatus = playing
    
    frmMain.tmrUpdateTime.Enabled = True
    
    If (Duration = 0) Then frmMain.tmrUpdateTime.Enabled = False
    
    SaveCurrentPos

End Sub

Public Sub Pause()
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    GlobalPlayStatus = Paused
    GlobalFilGraph.Pause
    UpdateStatus StaticString(PLAY_STATUS_PAUSED), PlayBack
    
    SaveCurrentPos

End Sub

Public Sub StopPlay()
    Precent = 0
    SaveCurrentPos
    GlobalPlayStatus = Stoped
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    
    If Not GlobalFilGraph Is Nothing Then
        GlobalFilGraph.Pause
        SaveCurrentPos

    End If
    
End Sub

Public Sub SwitchPlayStauts()
    
    If (mdlGlobalPlayer.Loaded) Then
        If (GlobalPlayStatus = playing) Then
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

Public Sub RaiseMediaFilter(list As vbalListViewCtl)
    list.ListItems.Add , , mdlGlobalPlayer.File

    Dim objFilter As IFilterInfo
    
    Dim objItem   As cListItem
    
    If (GlobalFilGraph Is Nothing) Then Exit Sub
    If (GlobalFilGraph.FilterCollection Is Nothing) Then Exit Sub
    
    For Each objFilter In GlobalFilGraph.FilterCollection
        
        list.ListItems.Add , , objFilter.Name
        
    Next
    
End Sub

Public Sub SwitchFullScreen(Optional force As Boolean = False, _
                            Optional forceValue As Boolean = False)

    If (Not HasVideo) Then Exit Sub
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
        UpdateTitle Dir(File)
        frmMain.WindowState = 2
    Else
    
        frmMain.BorderStyle = 2
        UpdateTitle Dir(File)
        frmMain.WindowState = 0

        If (lngStorageW <> 0) Then
            frmMain.Width = lngStorageW
            frmMain.Height = lngStorageH
            frmMain.Top = lngStorageT
            frmMain.Left = lngStorageL

        End If

    End If
    
End Sub

Public Sub SeekLastPos(ByVal strFullPath As String)
    mdlGlobalPlayer.CurrentTime = val(InI.INI_GetString(App.Path & "\LastPlayed.ini", "LastPos", MD5String(strFullPath)))
    
    If (mdlGlobalPlayer.CurrentTime = mdlGlobalPlayer.Duration) Then
        mdlGlobalPlayer.CurrentTime = 0
        
    End If
    
End Sub

Public Sub SeekCurrentPos()
    SeekLastPos mdlGlobalPlayer.File

End Sub

Public Sub SaveCurrentPos()

    If (mdlGlobalPlayer.Loaded) Then
        InI.INI_WriteString App.Path & "\LastPlayed.ini", "LastPos", MD5String(File), mdlGlobalPlayer.CurrentTime

    End If

End Sub
