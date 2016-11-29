Attribute VB_Name = "mdlStarup"
Option Explicit

Global IsIDE     As Boolean

Global IsRestart As Boolean

Private Declare Function RegisterSSubTmr6 _
                Lib "SSubTmr6.dll" _
                Alias "DllRegisterServer" () As Long
                
Private Declare Function RegistervbaListView6 _
                Lib "vbaListView6.ocx" _
                Alias "DllRegisterServer" () As Long

Private pAdminGroup       As IntPtr

Type SID_IDENTIFIER_AUTHORITY

    Value(6) As Byte

End Type

Private Const SECURITY_BUILTIN_DOMAIN_RID As Long = &H20

Private Const DOMAIN_ALIAS_RID_ADMINS     As Long = &H220

Private Const NULL_                       As Long = 0

Private Const SECURITY_NT_AUTHORITY       As Long = &H5

Private Declare Function AllocateAndInitializeSid _
                Lib "Advapi32.dll" (SID As SID_IDENTIFIER_AUTHORITY, _
                                    ByVal Count As Byte, _
                                    ByVal dwSubAuth0 As Long, _
                                    ByVal dwSubAuth1 As Long, _
                                    ByVal dwSubAuth2 As Long, _
                                    ByVal dwSubAuth3 As Long, _
                                    ByVal dwSubAuth4 As Long, _
                                    ByVal dwSubAuth5 As Long, _
                                    ByVal dwSubAuth6 As Long, _
                                    ByVal dwSubAuth7 As Long, _
                                    vpPSID As IntPtr) As BOOL
                                    
Private Declare Function CheckTokenMembership _
                Lib "Advapi32.dll" (ByVal hToken As IntPtr, _
                                    ByVal vpPSID As IntPtr, _
                                    ByRef isMember As BOOL) As BOOL

Private Declare Function FreeSid Lib "Advapi32.dll" (vpPSID As Long) As Long
                
Private Declare Function ShellExecuteEx _
                Lib "shell32" _
                Alias "ShellExecuteExW" (SEI As SHELLEXECUTEINFO2) As Long

Private Declare Function WaitForSingleObject _
                Lib "kernel32" (ByVal hHandle As Long, _
                                ByVal dwMilliseconds As Long) As Long

Private Const INFINITE = &HFFFF      '  Infinite timeout

Private Declare Function SetProcessDPIAware Lib "user32" () As BOOL

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As Long) As Long

Public isAdminPerm As Boolean

Public Sub RegisterCOM()
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\SSubTmr6.dll", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/u /s " & App.Path & "\vbaListView6.ocx", App.Path & "\", 0
    RegisterSSubTmr6
    RegistervbaListView6
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\SSubTmr6.dll", App.Path & "\", 0
    ShellExecute 0, "open", "regsvr32.exe", "/s " & App.Path & "\vbaListView6.ocx", App.Path & "\", 0
End Sub

Public Sub Main()

    IsIDE = GetIDEmode
    InitPerm
    
    InitConfigFiles

    'SetProcessDpiAwareness_soft
    If (IsSupportDIPSet) Then
        SetProcessDPIAware
        SetProcessDpiAwareness_hard

    End If

    If (Len(Command) = 8 And Left$(Command, 8) = "--regocx") Then
        RegisterCOM

        End

    End If

    If (Len(Command) = 6 And Left$(Command, 6) = "--perm") Then
        RegisterCOM

        End

    End If
    
    If (Len(Command) > 11 And Left$(Command, 9) = "--bindext") Then
        BindExt Mid(Command, 11)

        End
        
    End If

    If (Len(Command) > 13 And Left$(Command, 11) = "--unbindext") Then
        UnBindExt Mid(Command, 13)

        End
        
    End If
    
    If (Len(Command) = 11 And Left$(Command, 11) = "--unbindall") Then
        UnBindAll

        End
        
    End If

    If (Len(Command) = 13 And Left$(Command, 13) = "--association") Then
        AssociationRegister

        End
        
    End If

    If (Len(Command) = 11 And Left$(Command, 11) = "--uninstall") Then
        Uninstall

        End
        
    End If

    If (Len(Command) = 9 And Left$(Command, 9) = "--restart") Then
        IsRestart = True
        'GoTo FillDecoder

    End If
    
    If (App.PrevInstance) Then
        If (Len(Command) <> 0) Then

            Dim cbs    As COPYDATASTRUCT

            Dim ptrCmd As IntPtr

            ptrCmd = GetCommandLine()
            cbs.dwData = 3
            cbs.cbData = (lstrlen(ptrCmd) + 1) * 2
            cbs.lpData = ptrCmd
            SendMessageW val(GlobalConfig.LastHwnd), WM_COPYDATA, ByVal 0&, cbs

        End If
        
        'SendMessageW val(GlobalConfig.LastHwnd), PM_ACTIVE, 0&, 0&
        SetForegroundWindow GlobalConfig.LastHwnd
        
        End

        Exit Sub

    End If

    Load frmMenu

    If (GlobalConfig.Renderer = "") Then
        frmMenu.Renderers_Click RenderType.MadVRednerer
    Else
        frmMenu.Renderers_Click val(GlobalConfig.Renderer)

    End If

    On Error GoTo RegisterCOMErr

    mdlToolBarAlphaer.LoadUI
    
    If (Not IsIDE) Then StartHook frmMain.hWnd
    
    If (isAdminPerm) Then
        frmMain.Caption = "π‹¿Ì‘±: " & frmMain.Caption

    End If

    If (IsIDE) Then
        frmMain.Caption = "(IDE) " & frmMain.Caption

    End If

    On Error GoTo 0

    Dim i As Variant

    For Each i In GlobalConfig.LastPlayList

        If (i = "@") Then GoTo Placement
        mdlPlaylist.AddFileToPlaylist i
Placement:
    Next

    If (Not IsRestart And Not IsIDE) Then InitialCommandLine
    
    Exit Sub

RegisterCOMErr:
    'do gui com register
    ReqAdminPerm "--regocx"
    Shell App.Path & "\" & App.EXEName & ".exe --restart", vbNormalFocus

    End

    Resume

End Sub

Public Sub InitPerm()

    Dim isAdminPermission As BOOL
    
    Dim NtAuthority       As SID_IDENTIFIER_AUTHORITY
    
    NtAuthority.Value(5) = SECURITY_NT_AUTHORITY
    isAdminPermission = AllocateAndInitializeSid(NtAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, pAdminGroup)
    
    If (isAdminPermission) Then
        If (Not CheckTokenMembership(NULL_, pAdminGroup, isAdminPermission) = 1) Then
            CheckTokenMembership NULL_, pAdminGroup, isAdminPermission
            Debug.Print GetLastError
            isAdminPermission = False

        End If

        FreeSid pAdminGroup

    End If

    isAdminPerm = Not (isAdminPermission = 0)

End Sub

Public Function ReqAdminPerm(Optional strAction As String = "--perm")

    Dim sLInfo As SHELLEXECUTEINFO2

    With sLInfo
        .cbSize = Len(sLInfo)
        .lpVerb = StrPtr("runas")
        .lpFile = StrPtr(App.Path & "\" & App.EXEName & ".exe")
        .hWnd = 0
        .nShow = 1
        .lpParameters = StrPtr(strAction)
        
    End With

    If (ShellExecuteEx(sLInfo) = 0) Then
        
        MsgBox mdlLanguageApplyer.StaticString(BAD_PERMISSION_DENY)

        End

    End If
    
    WaitForSingleObject sLInfo.hProcess, INFINITE
    Sleep 1000
    
End Function

