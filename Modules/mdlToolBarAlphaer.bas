Attribute VB_Name = "mdlToolBarAlphaer"
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Private Const HTGROWBOX = 4

Private Const HTSIZE = HTGROWBOX

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public apMenuButton   As New AlphaPicture

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
    apMenuButton.LoadImageWH App.Path & "\Image\Menu.png", frmMain.bbMenuBar.height, frmMain.bbMenuBar.height
    
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
        apMenuButton.RefreshHW 32, 32
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
