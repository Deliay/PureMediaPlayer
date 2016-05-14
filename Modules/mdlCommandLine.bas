Attribute VB_Name = "mdlCommandLine"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private lPrevWndProc    As Long
Private hHookWindow     As Long
Public Const GWL_WNDPROC = -4
Public Const WM_DROPFILES = &H233
Public Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As Long)
Public Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileW" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As Long, ByVal ch As Long) As Long

Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As IntPtr
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As IntPtr, pNumArgs As Long) As IntPtr
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As IntPtr) As Long

Public Sub InitialCommandLine()
    Dim argc As Long, argv As IntPtr, i As Long
    argv = CommandLineToArgvW(GetCommandLine, argc)
    For i = 1 To argc - 1
        mdlPlaylist.AddFileToPlaylist AllocStr(ByVal PtrPtr(argv + vbPtrSize * i))
    Next
    If argc > 1 Then mdlPlaylist.PlayByName AllocStr(ByVal PtrPtr(argv + vbPtrSize * 1))
    LocalFree argv
End Sub

Public Sub SetHook(lHwnd As Long)
    
    If hHookWindow <> 0 Then Call Clearhook
    hHookWindow = lHwnd
    
    lPrevWndProc = SetWindowLong(hHookWindow, GWL_WNDPROC, AddressOf HookCallback)

End Sub

Public Sub Clearhook()

    Dim lReturn         As Long
    
    ' Check to be sure that there is a hook active
    If hHookWindow = 0 Then Exit Sub
    If IsEmpty(hHookWindow) = True Then Exit Sub
    If IsNull(hHookWindow) = True Then Exit Sub
    
    ' Remove the hook from the system
    lReturn = SetWindowLong(hHookWindow, GWL_WNDPROC, lPrevWndProc)

End Sub

Function HookCallback(ByVal hWnd As Long, _
            ByVal lMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    Select Case hWnd
        Case hHookWindow
            MessageProc lMsg, wParam, lParam
        Case Else
            ' The message is for some other window so ...
    
    End Select
    HookCallback = CallWindowProc(lPrevWndProc, hWnd, lMsg, wParam, lParam)

End Function

Public Sub MessageProc(lMsg As Long, _
                wParam As Long, _
                lParam As Long)
    Dim nDropCount          As Integer
    Dim nLoopCtr            As Integer
    Dim lReturn             As Long
    Dim hDrop               As Long
    Dim sFileName           As String
    Dim sShort              As String
    Dim sFirst              As String
    Select Case lMsg
        Case WM_DROPFILES
            hDrop = wParam
            sFileName = Space$(255)
            nDropCount = DragQueryFile(hDrop, -1, StrPtr(sFileName), 254)
            For nLoopCtr = 0 To nDropCount - 1
                sShort = Space$(255)
                sFileName = Space$(255)
                lReturn = DragQueryFile(hDrop, nLoopCtr, StrPtr(sFileName), 254)
                lReturn = GetShortPathName(StrPtr(sFileName), StrPtr(sShort), 254)
                If (Not (InStr(1, sFileName, Chr(0)) = 0)) Then
                    sFileName = Mid(sFileName, 1, InStr(1, sFileName, Chr(0)) - 1)
                End If
                mdlPlaylist.AddFileToPlaylist sFileName
                If (nLoopCtr = 0) Then sFirst = sFileName
            Next nLoopCtr
            Call DragFinish(hDrop)
            mdlPlaylist.PlayByName sFirst
    End Select
End Sub
