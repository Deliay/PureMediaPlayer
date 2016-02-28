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

Public apMenuButton   As New AlphaPicture
Public apPlaylistHint   As New AlphaPicture

Public UIHeightButtom As Long

Public UIHeightTop    As Long

Public UIWidthLeft    As Long

Public UIWidthRight   As Long

Private boolUIStatus  As Boolean

Public Sub LoadUI()
    Load frmMain
    Load frmPlaylist
    Load frmPaternAdd
    
    If (Dir(App.Path & "\language.ini") = "") Then
        CreateLanguagePart frmMenu
        CreateLanguagePart frmPlaylist
        CreateLanguagePart frmPaternAdd
    Else
        ApplyLanguageToForm frmMenu
        ApplyLanguageToForm frmPlaylist
        ApplyLanguageToForm frmPaternAdd

    End If
    
    frmPlaylist.Hide
    frmPaternAdd.Hide
    frmMain.Show
    apMenuButton.hDC = frmMain.bbMenuBar.hDC
    apMenuButton.LoadImageWH App.Path & "\Image\Menu.png", 32, 32
    apPlaylistHint.hDC = frmMain.bbPlaylist.hDC
    apPlaylistHint.LoadImageWH App.Path & "\Image\playlist_hint.png", 24, 48
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
        .bbMenuBar.Refresh
        .bbPlaylist.Visible = True
        .bbPlaylist.Refresh
        apMenuButton.RefreshHW 32, 32
        apPlaylistHint.RefreshHW 24, 48
        UIHeightButtom = UIHeightButtom + .pbTimeBar.height
        UIHeightButtom = UIHeightButtom + .sbStatusBar.height
        
        UIHeightTop = .bbMenuBar.height

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
