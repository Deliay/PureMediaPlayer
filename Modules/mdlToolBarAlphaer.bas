Attribute VB_Name = "mdlToolBarAlphaer"
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

'Public Const CFG_HISTORY_LAST_SAVE_PATH As String = "LastSavePath"
'
'Public Const CFG_HISTORY_LAST_OPEN_PATH As String = "LastOpenPath"
'
'Public Const CFG_SETTING_RENDERER       As String = "Renderer"
'
'Public Const CFG_SETTING_LANGUAGE       As String = "Language"
'
'Public Const CFG_SETTING_LAST_HWND      As String = "LastWindowHWND"

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

Type MD5_CTX

    dwNUMa      As Long
    dwNUMb      As Long
    Buffer(15)  As Byte
    cIN(63)     As Byte
    cDig(15)    As Byte

End Type

Private Declare Sub MD5Init Lib "advapi32" (lpContext As MD5_CTX)

Private Declare Sub MD5Final Lib "advapi32" (lpContext As MD5_CTX)

Private Declare Sub MD5Update _
                Lib "advapi32" (lpContext As MD5_CTX, _
                                ByRef lpBuffer As Any, _
                                ByVal BufSize As Long)
                                
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Private Declare Function GetDeviceCaps _
                Lib "gdi32.dll" (ByVal hDC As Long, _
                                 ByVal nIndex As Long) As Long

Private Declare Function ReleaseDC _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal hDC As Long) As Long

Const LOGPIXELSX   As Long = 8

Private stcContext As MD5_CTX

Public Function MD5String(strText As String) As String

    Dim aBuffer() As Byte
 
    Call MD5Init(stcContext)

    If (Len(strText) > 0) Then
        aBuffer = StrConv(strText, vbFromUnicode)
        Call MD5Update(stcContext, aBuffer(0), UBound(aBuffer) + 1)
    Else
        Call MD5Update(stcContext, 0, 0)

    End If

    Call MD5Final(stcContext)
    MD5String = stcContext.cDig
    
    Dim i&

    If (stcContext.dwNUMa = 0) Then
        MD5String = vbNullString
    Else
        MD5String = Space$(32)

        For i = 0 To 15
            Mid$(MD5String, i + i + 1) = Right$("0" & Hex$(stcContext.cDig(i)), 2)
        Next

    End If
   
End Function

Public Property Get UIStatus() As Boolean
    UIStatus = boolUIStatus

End Property

Public Sub LoadUI()
    Load frmMain
    Load frmPaternAdd

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
    
    mdlLanguageApplyer.EnumLanguageFile
    mdlLanguageApplyer.ReApplyLanguage

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
        'UIHeightButtom = UIHeightButtom + .pbTimeBar.Height * 2
        UIHeightButtom = UIHeightButtom + .sbStatusBar.Height
        
        UIHeightTop = 0

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
    frmMain.bbPlaystatus(PlayControl.CTRL_PLAYPAUSE).Cls
    apPlayControl(PlayControl.CTRL_PLAYPAUSE).hDC = frmMain.bbPlaystatus(PlayControl.CTRL_PLAYPAUSE).hDC

    If (mdlGlobalPlayer.GlobalPlayStatus = Running) Then
        
        apPlayControl(PlayControl.CTRL_PLAYPAUSE).LoadImageWH App.Path & "\Image\playcontrol_" & PlayControl.CTRL_PLAYPAUSE & "_.png", 32, 32
    Else
        apPlayControl(PlayControl.CTRL_PLAYPAUSE).LoadImageWH App.Path & "\Image\playcontrol_" & PlayControl.CTRL_PLAYPAUSE & ".png", 32, 32

    End If

End Sub
