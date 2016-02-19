Attribute VB_Name = "mdlGlobalPlayer"
Option Explicit

Public Const DefaultTitle As String = "Pure Media Player"

Public GlobalFilGraph     As FilgraphManager

Private BackupFilGraph    As New FilgraphManager

Private ifPostion         As IMediaPosition

Private ifPlayback        As IVideoWindow

Private ifVideo           As IBasicVideo

Private ifType            As IMediaTypeInfo

Public Width              As Long

Public Height             As Long

Private boolLoadedFile    As Boolean

Private boolIsFullScreen  As Boolean

Private strLastestFile    As String

Private VideoRatio        As Double

Public Enum StatusBarEnum

    Action = 1
    PlayBack = 2
    PlayTime = 3
    FileName = 4
    
End Enum

Public Enum PlayStatus

    Playing
    Paused
    Stoped

End Enum

Public GlobalPlayStatus As PlayStatus

Private Const WS_BORDER = &H800000

Private Const WS_CAPTION = &HC00000                                             '  WS_BORDER Or WS_DLGFRAME

Public Const WS_THICKFRAME = &H40000

Public Const WS_SIZEBOX = WS_THICKFRAME

Public Property Get Precent() As Single
    
    If ifPostion Is Nothing Then Exit Property
    Precent = (CurrentTime / Duration) * 100
    
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

Public Function NowFrame() As Picture
    
    Dim lngSize  As Long, lngHBITMAP As Long
    
    Dim buffer() As Long
    
    ifVideo.GetCurrentImage lngSize, 0&
    ReDim buffer(0 To lngSize)
    ifVideo.GetCurrentImage lngSize, buffer(0)
    lngHBITMAP = LongDIB2HBitmap(buffer)
    Set NowFrame = HBitmap2Picture(lngHBITMAP, 0)
    
End Function

Public Sub LoadMediaFile(strFilePath As String)
    UpdateStatus StaticString(PLAYER_STATUS_LOADING), Action
    Set GlobalFilGraph = BackupFilGraph
    'Set GlobalFilGraph = FilgraphManager
    strLastestFile = strFilePath
    UpdateStatus Dir(strFilePath), FileName
    frmPlayer.Caption = Dir(strFilePath)
    frmPlaylist.isHide = True
    frmPlaylist.Show vbModeless, frmMain
    frmPlaylist.AutoPatern
    frmPlaylist.AutoPatern
    frmPlaylist.File_PaternFind.Enabled = True
    
    On Error GoTo DcodeErr
    
    GlobalFilGraph.RenderFile strFilePath
    Set ifPostion = GlobalFilGraph
    
    On Error GoTo NotVideo
    Set ifVideo = GlobalFilGraph
NotVideo:
    Set ifPlayback = GlobalFilGraph
    
    On Error GoTo 0
    
    mdlGlobalPlayer.CurrentTime = 0
    mdlPlaylist.SetItemLength strFilePath, FormatedDuration
    
    If Not frmPlaylist.nowPlaying Is Nothing Then frmPlaylist.nowPlaying.Bold = False
    
    On Error GoTo notLoadPlaylist
    
    Set frmPlaylist.nowPlaying = frmPlaylist.lstPlaylist.ListItems(strFilePath)
notLoadPlaylist:
    
    If Not frmPlaylist.nowPlaying Is Nothing Then frmPlaylist.nowPlaying.Bold = True
    
    boolLoadedFile = True
    
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    
    On Error GoTo hErr
    
    ifPlayback.Caption = "PureMediaPlayer - LayerWindow"
    ifPlayback.Owner = frmPlayer.hWnd
    ifPlayback.MessageDrain = frmPlayer.hWnd
    
    Dim lngSrcStyle As Long
    
    lngSrcStyle = ifPlayback.WindowStyle
    lngSrcStyle = lngSrcStyle And Not WS_BORDER
    lngSrcStyle = lngSrcStyle And Not WS_CAPTION
    lngSrcStyle = lngSrcStyle And Not WS_SIZEBOX
    ifPlayback.WindowStyle = lngSrcStyle
hErr:
    mdlGlobalPlayer.Play
    DoEvents
    mdlGlobalPlayer.ResizePlayWindow
    
    ifPlayback.FullScreenMode = boolIsFullScreen
    
    Set BackupFilGraph = New FilgraphManager
    
    Exit Sub
DcodeErr:
    MsgBox "Not Support this codes type!"
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

Public Sub RaiseRegFileter(list As ListView)
    
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
    GlobalPlayStatus = Playing
    
    frmMain.tmrUpdateTime.Enabled = True
    
    If (Duration = 0) Then frmMain.tmrUpdateTime.Enabled = False
    
End Sub

Public Sub Pause()
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    GlobalPlayStatus = Paused
    GlobalFilGraph.Pause
    UpdateStatus StaticString(PLAY_STATUS_PAUSED), PlayBack
    
End Sub

Public Sub StopPlay()
    Precent = 0
    GlobalPlayStatus = Stoped
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    
    If Not GlobalFilGraph Is Nothing Then GlobalFilGraph.Stop
    
End Sub

Public Sub SwitchPlayStauts()
    
    If (mdlGlobalPlayer.Loaded) Then
        If (GlobalPlayStatus = Playing) Then
            mdlGlobalPlayer.Pause
        ElseIf (GlobalPlayStatus = Paused) Then
            mdlGlobalPlayer.Play
            
        End If
        
    End If
    
End Sub

Public Sub CloseFile()
    UpdateStatus StaticString(PLAYER_STATUS_CLOSEING), Action
    strLastestFile = ""
    frmPlayer.Caption = StaticString(PLAYER_STATUS_IDLE)
    DoEvents
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    
End Sub

Public Sub ResizePlayWindow()
    
    If (ifVideo Is Nothing) Then Exit Sub
    
    On Error GoTo hErr
    
    Dim commonW As Long, commonH As Long
    
    Dim resultW As Long, resultH As Long
    
    Dim resultT As Long, resultL As Long
    
    ifVideo.GetVideoSize commonW, commonH
    VideoRatio = commonW / commonH
    resultW = Width
    resultH = Width / VideoRatio
    
    If (resultH > Height) Then
        resultW = VideoRatio * Height
        resultH = Height
        
    End If
    
    resultT = (Height - resultH) / 2
    resultL = (Width - resultW) / 2
    ifPlayback.SetWindowPosition resultL, resultT, resultW, resultH
hErr:
    frmMain.pbTimeBar.ZOrder 0
    
End Sub

Public Sub UpdateStatus(strCaption As String, Target As StatusBarEnum)
    frmMain.sbStatusBar.Panels.Item(CLng(Target) * 2 - 1).Text = strCaption
    
End Sub

Public Sub RaiseMediaFilter(list As ListView)
    
    Dim objFilter As IFilterInfo
    
    Dim objItem   As ListItem
    
    If (GlobalFilGraph Is Nothing) Then Exit Sub
    If (GlobalFilGraph.FilterCollection Is Nothing) Then Exit Sub
    
    For Each objFilter In GlobalFilGraph.FilterCollection
        
        list.ListItems.Add(, , objFilter.Name).SubItems(1) = objFilter.VendorInfo
        
    Next
    
End Sub

Public Sub ExitProgram()
    StopPlay
    CloseFile
    End
    
End Sub

Public Sub SwitchFullScreen(Optional force As Boolean = False, _
                            Optional forceValue As Boolean = False)
    
    If (ifPlayback Is Nothing) Then Exit Sub
    If force = True Then
        boolIsFullScreen = forceValue
        ifPlayback.FullScreenMode = forceValue
        Exit Sub
        
    End If
    
    boolIsFullScreen = Not boolIsFullScreen
    ifPlayback.FullScreenMode = boolIsFullScreen
    
End Sub
