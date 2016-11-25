Attribute VB_Name = "mdlMsgProc"
Option Explicit

Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongW" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function CallAsmCode _
                Lib "user32" _
                Alias "CallWindowProcA" (lpPrevWndFunc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         lParam As Long) As Long

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
                Lib "shell32.dll" (ByVal hWnd As Long, _
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
                Alias "PostMessageW" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetWindowThreadProcessId _
                Lib "user32" (ByVal hWnd As Long, _
                              lpdwProcessId As Long) As Long

Private Declare Function AttachThreadInput _
                Lib "user32" (ByVal idAttach As Long, _
                              ByVal idAttachTo As Long, _
                              ByVal fAttach As Long) As Long

Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function lstrlen _
               Lib "kernel32" _
               Alias "lstrlenW" (ByVal lpString As Long) As Long

Public Declare Sub CopyMemory _
               Lib "kernel32" _
               Alias "RtlMoveMemory" (hpvDest As Any, _
                                      hpvSource As Any, _
                                      ByVal cbCopy As Long)

Public Const MAX_PATH = 260

Public Sub ReadDrapQueryFile(ByVal hDrop As Long)

    Dim i As Long, Count As Long, buf As String, sbuf As String

    Count = DragQueryFile(hDrop, i, 0&, 0&)
    
    For i = 0 To Count - 1
        buf = Space$(255)
        sbuf = Space$(255)
        DragQueryFile hDrop, i, StrPtr(buf), 254
        GetShortPathName StrPtr(buf), StrPtr(sbuf), 255

        If (Not (InStr(1, sbuf, Chr(0)) = 0)) Then
            sbuf = Mid(sbuf, 1, InStr(1, sbuf, Chr(0)) - 1)

        End If
        
        Select Case LCase((Right$(sbuf, 3)))

            Case "idx", "sub", "srt", "ssa", "smi", "ssa", "ass", "sup"

                If (mdlGlobalPlayer.GlobalPlayStatus = Running) Then mdlFilterProductor.SetVSFilterFileName sbuf

                Exit Sub
            
        End Select

        mdlPlaylist.AddFileToPlaylist sbuf
    Next

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

                Select Case LCase((Right$(sFileName, 3)))

                    Case "idx", "sub", "srt", "ssa", "smi", "ssa", "ass", "sup"

                        If (mdlGlobalPlayer.GlobalPlayStatus = Running) Then mdlFilterProductor.SetVSFilterFileName sFileName

                        Exit Sub
                    
                End Select

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
            SetForegroundWindow frmMain.hWnd
            BringWindowToTop frmMain.hWnd
            AttachThreadInput bPid, nPid, False
            frmMain.SetFocus

    End Select

End Sub

Private Function CallAnyFunc(ByVal pFn As Long, _
                             ByVal pParam As Long, _
                             ByVal Count As Long) As Long

    Dim CallAnyFuncCode(34) As Long, lRet As Long

    CallAnyFuncCode(0) = &H53EC8B55
    CallAnyFuncCode(1) = &HE8&
    CallAnyFuncCode(2) = &HEB815B00
    CallAnyFuncCode(3) = &H1000112C
    CallAnyFuncCode(4) = &H114A938D
    CallAnyFuncCode(5) = &H64521000
    CallAnyFuncCode(6) = &H35FF&
    CallAnyFuncCode(7) = &H89640000
    CallAnyFuncCode(8) = &H25&
    CallAnyFuncCode(9) = &H8B1FEB00
    CallAnyFuncCode(10) = &HE80C2444
    CallAnyFuncCode(11) = &H0&
    CallAnyFuncCode(12) = &H53E98159
    CallAnyFuncCode(13) = &H8D100011
    CallAnyFuncCode(14) = &H119791
    CallAnyFuncCode(15) = &HB8908910
    CallAnyFuncCode(16) = &H33000000
    CallAnyFuncCode(17) = &H558BC3C0
    CallAnyFuncCode(18) = &H104D8B0C
    CallAnyFuncCode(19) = &HEB8A148D
    CallAnyFuncCode(20) = &HFC528D06
    CallAnyFuncCode(21) = &HB4932FF
    CallAnyFuncCode(22) = &H8BF675C9
    CallAnyFuncCode(23) = &HD0FF0845
    CallAnyFuncCode(24) = &H58F64
    CallAnyFuncCode(25) = &H83000000
    CallAnyFuncCode(26) = &H4D8B04C4
    CallAnyFuncCode(27) = &H5B018914
    CallAnyFuncCode(28) = &H10C2C9
    CallAnyFuncCode(29) = &H58F64
    CallAnyFuncCode(30) = &H83000000
    CallAnyFuncCode(31) = &HC03304C4
    CallAnyFuncCode(32) = &H89144D8B
    CallAnyFuncCode(33) = &HC2C95B21
    CallAnyFuncCode(34) = &H90900010
    CallAnyFunc = CallAsmCode(CallAnyFuncCode(0), pFn, pParam, Count, lRet)

    If CallAnyFunc <> lRet Then
        CallAnyFunc = 0 '??????????,???????????????
        Debug.Assert False '??????????,?????????????

    End If

End Function

