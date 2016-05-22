Attribute VB_Name = "mdlCommandLine"
Option Explicit

Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongW" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Type COPYDATASTRUCT

    dwData As Long
    cbData As Long
    lpData As Long

End Type

Public Const GWL_WNDPROC = (-4)

Public Const WM_COPYDATA = &H4A

Private lPrevWndProc As Long

Private hHookWindow  As Long

Private Const WM_DROPFILES = &H233

Private Const WM_HOTKEY = &H312

Private Const WM_CLOSE = &H10

'wParam
Public Const PM_ADDMEDIAFILE = &HFFF

Public Const PM_ACTIVE = &HFFE

Public Const PM_PLAY_LAST = &HFFD

Private Declare Sub DragAcceptFiles _
                Lib "shell32.dll" (ByVal hwnd As Long, _
                                   ByVal fAccept As Long)

Private Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)

Private Declare Function DragQueryFile _
                Lib "shell32.dll" _
                Alias "DragQueryFileW" (ByVal hDrop As Long, _
                                        ByVal UINT As Long, _
                                        ByVal lpStr As Long, _
                                        ByVal ch As Long) As Long

Public Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As IntPtr

Public Declare Function CommandLineToArgvW _
               Lib "shell32" (ByVal lpCmdLine As IntPtr, _
                              pNumArgs As Long) As IntPtr

Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As IntPtr) As Long

Private Declare Function PostMessage _
                Lib "user32" _
                Alias "PostMessageW" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Declare Function TerminateProcess _
                Lib "kernel32" (ByVal hProcess As Long, _
                                ByVal uExitCode As Long) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hwnd As Long, _
                              lpdwProcessId As Long) As Long

Private Declare Function AttachThreadInput _
                Lib "user32" (ByVal idAttach As Long, _
                              ByVal idAttachTo As Long, _
                              ByVal fAttach As Long) As Long

Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function lstrlen _
               Lib "kernel32" _
               Alias "lstrlenW" (ByVal lpString As Long) As Long

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (hpvDest As Any, _
                                      hpvSource As Any, _
                                      ByVal cbCopy As Long)
         
Public Sub InitialCommandLine()

    Dim argc As Long, argv As IntPtr, i As Long

    argv = CommandLineToArgvW(GetCommandLine, argc)
    
    For i = 1 To argc - 1
        mdlPlaylist.AddFileToPlaylist AllocStr(ByVal PtrPtr(argv + vbPtrSize * i))
    Next

    If argc > 1 Then mdlPlaylist.PlayByName AllocStr(ByVal PtrPtr(argv + vbPtrSize * 1))
    LocalFree argv

End Sub

Private Sub SetHook(lHwnd As Long)
    
    If hHookWindow <> 0 Then Call Clearhook
    hHookWindow = lHwnd
    
    lPrevWndProc = SetWindowLong(hHookWindow, GWL_WNDPROC, AddressOf HookCallback)

End Sub

Public Sub StartHook(lHwnd As Long)
    SetHook (lHwnd)
    
    DragAcceptFiles (lHwnd), True

End Sub

Public Sub Clearhook()

    Dim lReturn As Long
    
    ' Check to be sure that there is a hook active
    If hHookWindow = 0 Then Exit Sub
    If IsEmpty(hHookWindow) = True Then Exit Sub
    If IsNull(hHookWindow) = True Then Exit Sub
    
    ' Remove the hook from the system
    lReturn = SetWindowLong(hHookWindow, GWL_WNDPROC, lPrevWndProc)

End Sub

Function HookCallback(ByVal hwnd As Long, _
                      ByVal lMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long) As Long

    Select Case hwnd

        Case hHookWindow
            MessageProc lMsg, wParam, lParam

        Case Else
            ' The message is for some other window so ...
    
    End Select

    HookCallback = CallWindowProc(lPrevWndProc, hwnd, lMsg, wParam, lParam)

End Function

Public Sub MessageProc(lMsg As Long, wParam As Long, lParam As Long)

    Dim nDropCount As Integer

    Dim nLoopCtr   As Integer

    Dim lReturn    As Long

    Dim hDrop      As Long

    Dim sFileName  As String

    Dim sShort     As String

    Dim sFirst     As String

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

        Case WM_CLOSE
            ExitProgram
            
        Case PM_ADDMEDIAFILE

            'ABORT FOR USE
        Case WM_COPYDATA

            Dim cbs As COPYDATASTRUCT

            CopyMemory cbs, ByVal lParam, Len(cbs)

            Dim argc      As Long, argv As IntPtr, i As Long

            Dim strResult As String

            strResult = Space$(cbs.cbData)
            CopyMemory ByVal StrPtr(strResult), ByVal cbs.lpData, cbs.cbData
            argv = CommandLineToArgvW(cbs.lpData, argc)

            For i = 1 To argc - 1
                mdlPlaylist.AddFileToPlaylist AllocStr(ByVal PtrPtr(argv + vbPtrSize * i))
            Next
        
            If argc > 1 Then mdlPlaylist.PlayByName AllocStr(ByVal PtrPtr(argv + vbPtrSize * 1))
            LocalFree argv
            
        Case PM_PLAY_LAST
            mdlPlaylist.PlayByName mdlPlaylist.colPlayItems(mdlPlaylist.colPlayItems.Count).FullPath
            
        Case PM_ACTIVE

            Dim bPid As Long, nPid As Long

            bPid = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
            nPid = GetCurrentProcessId()
            AttachThreadInput bPid, nPid, True
            SetForegroundWindow frmMain.hwnd
            BringWindowToTop frmMain.hwnd
            AttachThreadInput bPid, nPid, False
            frmMain.SetFocus

    End Select

End Sub

Public Sub ExitProgram()

    If (mdlGlobalPlayer.Loaded) Then
        InI.INI_WriteString App.Path & "\LastPlayed.ini", "LastPos", mdlGlobalPlayer.File, mdlGlobalPlayer.CurrentTime

    End If

    StopPlay
    CloseFile
    Clearhook
    TerminateProcess GetCurrentProcessId, 5
    Shell "taskkill /F /PID " & GetCurrentProcessId
    End

End Sub
