Attribute VB_Name = "mdlStartupHelper"
Option Explicit

Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function CloseClipboard Lib "user32" () As Long

Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long

Private Declare Function GetModuleFileName _
                Lib "kernel32" _
                Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                            ByVal lpFileName As String, _
                                            ByVal nSize As Long) As Long

Private Declare Function LoadLibrary _
                Lib "kernel32" _
                Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress _
                Lib "kernel32" (ByVal hModule As Long, _
                                ByVal lpProcName As String) As Long

Private Declare Function SetProcessDpiAwareness _
                Lib "Shcore.dll" (ByVal Value As Long) As Long

Private Declare Function TerminateProcess _
                Lib "kernel32" (ByVal hProcess As Long, _
                                ByVal uExitCode As Long) As Long

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Function SetProcessDpiAwareness_hard()
    SetProcessDpiAwareness 2&

End Function

Public Function IsSupportDIPSet() As Boolean

    Dim lLib As Long, hAddress As Long
    
    lLib = LoadLibrary("Shcore.dll")
    IsSupportDIPSet = False

    If (lLib <> 0) Then
        hAddress = GetProcAddress(lLib, "SetProcessDpiAwareness")

        If (hAddress <> 0) Then
            IsSupportDIPSet = True
            Exit Function

        End If

    End If

End Function

Public Function GetIDEmode() As Boolean

    Dim strFileName As String

    Dim lngCount    As Long

    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left(strFileName, lngCount)

    If NameGet(strFileName) = "VB6.EXE" Then

        GetIDEmode = True
    Else

        GetIDEmode = False

    End If

End Function

Public Sub InitialCommandLine()

    Dim argc As Long, argv As IntPtr, i As Long

    argv = CommandLineToArgvW(GetCommandLine, argc)
    
    For i = 1 To argc - 1
        mdlPlaylist.AddFileToPlaylist AllocStr(ByVal PtrPtr(argv + vbPtrSize * i))
    Next

    If argc > 1 Then mdlPlaylist.PlayByName AllocStr(ByVal PtrPtr(argv + vbPtrSize * 1))
    LocalFree argv

End Sub

Public Sub ExitProgram()

    If (mdlGlobalPlayer.Loaded) Then
        GlobalConfig.LastPlayPos.Value(mdlGlobalPlayer.FileMD5) = CStr(mdlGlobalPlayer.CurrentTime)

    End If
    
    mdlConfig.SaveConfig
    
    StopPlay
    CloseFile
    Clearhook

    If (Not IsIDE) Then TerminateProcess GetCurrentProcessId, 5
    If (Not IsIDE) Then Shell "taskkill /F /PID " & GetCurrentProcessId
    End

End Sub

