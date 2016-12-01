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

Private Declare Sub MD5Init Lib "advapi32" (lpContext As MD5_CTX)

Private Declare Sub MD5Final Lib "advapi32" (lpContext As MD5_CTX)

Private Declare Sub MD5Update _
                Lib "advapi32" (lpContext As MD5_CTX, _
                                ByRef lpBuffer As Any, _
                                ByVal BufSize As Long)
                                
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private stcContext As MD5_CTX

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

Public Sub AssociationRegister()

    Dim reg         As New RegisterEditor

    Dim strCurrPath As String

    strCurrPath = "Applications\PureMediaPlayer.exe"
    
    '1. perhap the Application Register
    If Not (reg.ItemExits(HKEY_CLASSES_ROOT, strCurrPath)) Then
    
        If Not (reg.CreateKey(HKEY_CLASSES_ROOT, strCurrPath)) Then
            MsgBox "Fail on Create Application Key"
            
        End If
        
    End If
    
    strCurrPath = strCurrPath & "\shell\open\Command"
    reg.SetString HKEY_CLASSES_ROOT, strCurrPath, "", """" & App.Path & "\" & App.EXEName & ".exe"" " & """%1"""
    
    '2. ready the class
    If Not (reg.ItemExits(HKEY_CLASSES_ROOT, "PureMediaPlayer")) Then
    
        If Not (reg.CreateKey(HKEY_CLASSES_ROOT, "PureMediaPlayer")) Then
            MsgBox "Fail on Create ProgID"
            
        End If
        
    End If
    
    strCurrPath = "PureMediaPlayer\shell\open\Command"
    reg.SetString HKEY_CLASSES_ROOT, strCurrPath, "", """" & App.Path & "\" & App.EXEName & ".exe"" " & """%1"""
    
    GlobalConfig.AppRegistered = "1"

    mdlConfig.SaveConfig
    
End Sub

Public Sub BindExt(ByVal strExt As String)

    Dim reg As New RegisterEditor
    
    If Not (reg.ItemExits(HKEY_CLASSES_ROOT, strExt)) Then
    
        If Not (reg.CreateKey(HKEY_CLASSES_ROOT, strExt)) Then
            MsgBox "Fail on Create System Associations"
            
        End If
        
    End If
    
    reg.SetString HKEY_CLASSES_ROOT, strExt, "", "PureMediaPlayer"
    
    MsgBox mdlLanguageApplyer.StaticString(EXT_BIND_SUCCESS)
    
End Sub

Public Sub Uninstall()

    Dim reg As New RegisterEditor
    
    reg.DelKey HKEY_CLASSES_ROOT, "PureMediaPlayer"
    reg.DelKey HKEY_CLASSES_ROOT, "Applications\PureMediaPlayer.exe"
    
    'remove main key
    UnBindAll
    
End Sub

Public Sub UnBindAll()

    Dim i As Variant, k As String

    GlobalConfig.BindedFileExts.Remove 1
    
    For Each i In GlobalConfig.BindedFileExts

        k = i
        UnBindExt k
        
    Next
    
    GlobalConfig.BindedFileExts.Clear

End Sub

Public Sub UnBindExt(ByVal strExt As String)

    If (GlobalConfig.OldBindExts.Exist(strExt)) Then

        'exist a old value/setting
        Dim reg As New RegisterEditor

        reg.SetString HKEY_CLASSES_ROOT, strExt, "", GlobalConfig.OldBindExts(strExt)
        
    Else
    
        If (reg.GetString(HKEY_CLASSES_ROOT, strExt, "") = "PureMediaPlayer") Then
            reg.DelKey HKEY_CLASSES_ROOT, strExt
            
        End If
        
    End If
    
End Sub

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
