Attribute VB_Name = "mdlLanguageApplyer"
Option Explicit

Private Declare Function MultiByteToWideChar _
                Lib "kernel32.dll" (ByVal CodePage As Long, _
                                    ByVal dwFlags As Long, _
                                    ByVal lpMultiByteStr As String, _
                                    ByVal cchMultiByte As Long, _
                                    ByVal lpWideCharStr As String, _
                                    ByVal cchWideChar As Long) As Long

Private Declare Function WideCharToMultiByte _
                Lib "kernel32.dll" (ByVal CodePage As Long, _
                                    ByVal dwFlags As Long, _
                                    ByVal lpWideCharStr As Long, _
                                    ByVal cchWideChar As Long, _
                                    ByRef lpMultiByteStr As Any, _
                                    ByVal cchMultiByte As Long, _
                                    ByVal lpDefaultChar As String, _
                                    ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_ACP  As Long = 0

Private Const CP_UTF8 As Long = 65001

Public Enum STATIC_STRING_ENUM

    PLAY_STATUS_PLAYING
    PLAY_STATUS_PAUSED
    PLAY_STATUS_STOPED
        
    PLAYER_STATUS_IDLE
    PLAYER_STATUS_LOADING
    PLAYER_STATUS_READY
    PLAYER_STATUS_CLOSEING
        
    FILE_STATUS_NOFILE

End Enum

Const DEFAULT_PLAY_STATUS_PLAYING    As String = "Playing"

Const DEFAULT_PLAY_STATUS_PAUSED     As String = "Paused"

Const DEFAULT_PLAY_STATUS_STOPED     As String = "Stoped"

Const DEFAULT_PLAYER_STATUS_IDLE     As String = "Idle"

Const DEFAULT_PLAYER_STATUS_LOADING  As String = "Loading"

Const DEFAULT_PLAYER_STATUS_READY    As String = "Ready"

Const DEFAULT_PLAYER_STATUS_CLOSEING As String = "Closeing"

Const DEFAULT_FILE_STATUS_NOFILE     As String = "No File Opend"

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Public Function DefaultStaticString(ctype As STATIC_STRING_ENUM) As String

    Select Case ctype

        Case STATIC_STRING_ENUM.PLAY_STATUS_PAUSED
            DefaultStaticString = DEFAULT_PLAY_STATUS_PAUSED

        Case STATIC_STRING_ENUM.PLAY_STATUS_PLAYING
            DefaultStaticString = DEFAULT_PLAY_STATUS_PLAYING

        Case STATIC_STRING_ENUM.PLAY_STATUS_STOPED
            DefaultStaticString = DEFAULT_PLAY_STATUS_STOPED

        Case STATIC_STRING_ENUM.PLAYER_STATUS_CLOSEING
            DefaultStaticString = DEFAULT_PLAYER_STATUS_CLOSEING

        Case STATIC_STRING_ENUM.PLAYER_STATUS_IDLE
            DefaultStaticString = DEFAULT_PLAYER_STATUS_IDLE

        Case STATIC_STRING_ENUM.PLAYER_STATUS_LOADING
            DefaultStaticString = DEFAULT_PLAYER_STATUS_LOADING

        Case STATIC_STRING_ENUM.PLAYER_STATUS_READY
            DefaultStaticString = DEFAULT_PLAYER_STATUS_READY
            
        Case STATIC_STRING_ENUM.FILE_STATUS_NOFILE
            DefaultStaticString = DEFAULT_FILE_STATUS_NOFILE

    End Select

End Function

Public Sub ApplyLanguageToForm(frm As Form)

    Dim objCtrl As Control
    Dim boolHaveCaption As Boolean
    For Each objCtrl In frm.Controls

        On Error GoTo Continue

        Dim strVal As String
        boolHaveCaption = True
        strVal = objCtrl.Caption
        If (boolHaveCaption = False) Then GoTo Continue
        If (Len(strVal) > 0) Then
            If (strVal <> "-") Then
                strVal = GetLanguage(frm.Name, objCtrl.Name)
                If (strVal = "") Then
                    InI.INI_WriteString App.Path & "\language.ini", frm.Name, objCtrl.Name, objCtrl.Caption
                Else
                    objCtrl.Caption = strVal
                End If
            Else
                GoTo Continue
            End If
        End If
Continue:
        boolHaveCaption = False
        Resume Next
    Next

End Sub

Public Function GetLanguage(strPart As String, strKey As String) As String
    GetLanguage = InI.INI_GetString(App.Path & "\language.ini", strPart, strKey)

End Function

Public Sub CreateLanguagePart(frm As Form)

    Dim objCtrl As Control

    On Error Resume Next
    
    For Each objCtrl In frm.Controls

        Dim strVal As String

        strVal = objCtrl.Caption

        If (Len(strVal) <> 0) Then
            If (strVal <> "-") Then InI.INI_WriteString App.Path & "\language.ini", frm.Name, objCtrl.Name, strVal

        End If

    Next

End Sub

Public Function StaticString(ctype As STATIC_STRING_ENUM)
    StaticString = InI.INI_GetString(App.Path & "\language.ini", "StaticString", "String" & ctype)

    If (Len(StaticString) = 0) Then
        InI.INI_WriteString App.Path & "\language.ini", "StaticString", "String" & ctype, DefaultStaticString(ctype)
        StaticString = DefaultStaticString(ctype)
    
    End If

End Function
