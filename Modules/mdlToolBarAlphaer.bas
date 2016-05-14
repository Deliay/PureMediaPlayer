Attribute VB_Name = "mdlToolBarAlphaer"
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Private Const HTGROWBOX = 4

Private Const HTSIZE = HTGROWBOX

Public Const RGN_OR = 2

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
                                        
Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long

Private Declare Function SetWindowRgn _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal hRgn As Long, _
                              ByVal bRedraw As Boolean) As Long

Private Declare Function CreateRectRgn _
                Lib "gdi32" (ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long) As Long

Private Declare Function CombineRgn _
                Lib "gdi32" (ByVal hDestRgn As Long, _
                             ByVal hSrcRgn1 As Long, _
                             ByVal hSrcRgn2 As Long, _
                             ByVal nCombineMode As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Enum PlayControl

    CTRL_PLAYPAUSE
    CTRL_STOP
    CTRL_NEXT
    CTRL_PREV
    CTRL_VOICE

End Enum

Public apMenuButton       As New AlphaPicture

Public apPlaylistHint     As New AlphaPicture

Public apPlayControl(4)   As New AlphaPicture

Public UIHeightButtom     As Long

Public UIHeightTop        As Long

Public UIWidthLeft        As Long

Public UIWidthRight       As Long

Private boolUIStatus      As Boolean

Public boolPlaylistStatus As Boolean

Public Property Get UIStatus() As Boolean
    UIStatus = boolUIStatus
End Property

Public Sub LoadUI()
    Load frmMain
    Load frmPaternAdd

    If (Dir(App.Path & "\language.ini") = "") Then
        CreateLanguagePart frmMenu
        CreateLanguagePart frmPaternAdd
    Else
        ApplyLanguageToForm frmMenu
        ApplyLanguageToForm frmPaternAdd

    End If
    
    frmPaternAdd.Hide
    
    apMenuButton.hDC = frmMain.bbMenuBar.hDC
    apMenuButton.LoadImageWH App.Path & "\Image\Menu.png", 32, 32
    apPlaylistHint.hDC = frmMain.bbPlaylist.hDC
    apPlaylistHint.LoadImageWH App.Path & "\Image\playlist_hint.png", 24, 48
    
    Dim i As Long

    For i = 0 To 4
        apPlayControl(i).hDC = frmMain.bbPlaystatus(i).hDC
        apPlayControl(i).LoadImageWH App.Path & "\Image\playcontrol_" & i & ".png", 32, 32
    Next
    frmMain.Show
End Sub

Public Sub RefreshUI()

    With frmMain
        UIHeightButtom = 0
        UIHeightTop = 0
        UIWidthLeft = 0
        UIWidthRight = 0
        .sbStatusBar.Visible = True
        .pbTimeBar.Visible = True
        .bbMenuBar.Visible = True
        .bbMenuBar.ZOrder 0
        .bbPlaylist.Visible = True
        .bbPlaylist.ZOrder 0
        UIHeightButtom = UIHeightButtom + .pbTimeBar.Height
        UIHeightButtom = UIHeightButtom + .sbStatusBar.Height
        
        UIHeightTop = .bbMenuBar.Height
        Dim i As Long
    
        For i = 0 To 4
            .bbPlaystatus(i).Visible = True
        Next
        
    End With

End Sub

Public Sub HideUI()

    With frmMain
        UIHeightButtom = 0
        UIHeightTop = 0
        UIWidthLeft = 0
        UIWidthRight = 0
        .pbTimeBar.Visible = False
        .sbStatusBar.Visible = False
        .bbMenuBar.Visible = False
        .bbPlaylist.Visible = False

        Dim i As Long

        For i = 0 To 4
            .bbPlaystatus(i).Visible = False
        Next

    End With
    
End Sub

Public Sub SwitchUI(Optional force As Boolean = False, Optional val As Boolean = True)

    If (force = True) Then
        boolUIStatus = val
        GoTo refreshDirect

    End If

    boolUIStatus = Not boolUIStatus
    
refreshDirect:

    If (boolUIStatus) Then
        RefreshUI
        
    Else
        HideUI
        
    End If
    frmMain.Form_Resize
    frmMain.srcH = 0
End Sub

Public Sub DragWindow(ByVal lngHwnd As Long)
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub

Public Sub SizeWindow(ByVal lngHwnd As Long)
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTSIZE, 0&

End Sub

Public Sub NoBorder(ByVal lngHwnd As Long)
    SetWindowLong lngHwnd, (-16), &H80000000 Or &H20000 Or &H80000 Or &H10000000
    
End Sub

Public Sub PlaylistShow()
    mdlToolBarAlphaer.UIWidthRight = frmMain.lstPlaylist.Width
    frmMain.lstPlaylist.Left = (frmMain.Width / Screen.TwipsPerPixelX) - frmMain.lstPlaylist.Width
    frmMain.ReCalcPlayWindow
    boolPlaylistStatus = True

End Sub

Public Sub PlaylistHide()
    mdlToolBarAlphaer.UIWidthRight = 0
    frmMain.lstPlaylist.Left = (frmMain.Width / Screen.TwipsPerPixelX) + frmMain.lstPlaylist.Width
    frmMain.ReCalcPlayWindow
    boolPlaylistStatus = False

End Sub

Public Sub PlayPauseSwitch()
    If (mdlGlobalPlayer.GlobalPlayStatus = playing) Then
        apPlayControl(PlayControl.CTRL_PLAYPAUSE).LoadImageWH App.Path & "\Image\playcontrol_" & PlayControl.CTRL_PLAYPAUSE & "_.png", 32, 32
    Else
        apPlayControl(PlayControl.CTRL_PLAYPAUSE).LoadImageWH App.Path & "\Image\playcontrol_" & PlayControl.CTRL_PLAYPAUSE & ".png", 32, 32

    End If

End Sub
