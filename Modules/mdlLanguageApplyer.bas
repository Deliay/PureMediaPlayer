Attribute VB_Name = "mdlLanguageApplyer"
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Enum STATIC_STRING_ENUM

    PLAY_STATUS_PLAYING
    PLAY_STATUS_PAUSED
    PLAY_STATUS_STOPED
        
    PLAYER_STATUS_IDLE
    PLAYER_STATUS_LOADING
    PLAYER_STATUS_READY
    PLAYER_STATUS_CLOSEING

    FILE_STATUS_NOFILE

    FILE_NOT_SUPPORT
    PLAYER_VIOCE_RATE

    TIPS_ALREADY_RUN

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
                                       
Private colLanguages As New Collection

Public LanguageIndex As Long

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

    Dim objCtrl         As Control

    Dim boolHaveCaption As Boolean

    For Each objCtrl In frm.Controls
        
        If (InStr(1, objCtrl.Name, "Language") <> 0) Then
            GoTo NextElement

        End If
        
        On Error GoTo Continue

        Dim strVal As String

        boolHaveCaption = True
        strVal = objCtrl.Caption

        If (boolHaveCaption = False) Then GoTo Continue
        If (Len(strVal) > 0) Then
            If (strVal <> "-") Then
                strVal = GetLanguage(frm.Name, objCtrl.Name & objCtrl.Index)

                If (strVal = "") Then
                    InI.INI_WriteString GetLanguageFile, frm.Name, objCtrl.Name & objCtrl.Index, objCtrl.Caption
                Else
                    objCtrl.Caption = strVal

                End If

            Else
                GoTo Continue

            End If

        End If

Continue:
        boolHaveCaption = False
        strVal = GetLanguage(frm.Name, objCtrl.Name)

        Resume Next

NextElement:
    Next

End Sub

Public Function GetLanguage(strPart As String, strKey As String) As String
    GetLanguage = InI.INI_GetString(GetLanguageFile, strPart, strKey)

End Function

Public Sub CreateLanguagePart(frm As Form)

    Dim objCtrl As Control

    On Error Resume Next
    
    For Each objCtrl In frm.Controls

        Dim strVal As String

        strVal = objCtrl.Caption

        If (Len(strVal) <> 0) Then
            If (strVal <> "-") Then InI.INI_WriteString GetLanguageFile, frm.Name, objCtrl.Name & objCtrl.Index, strVal

        End If

    Next

End Sub

Public Function StaticString(ctype As STATIC_STRING_ENUM)
    StaticString = InI.INI_GetString(GetLanguageFile, "StaticString", "String" & ctype)

    If (Len(StaticString) = 0) Then
        InI.INI_WriteString GetLanguageFile, "StaticString", "String" & ctype, DefaultStaticString(ctype)
        StaticString = DefaultStaticString(ctype)
    
    End If

End Function

Public Function GetLanguageName() As String
    GetLanguageName = InI.INI_GetString(GetLanguageFile, "Lang", "ShowName")

End Function

Public Function GetLanguageFile() As String

    GetLanguageFile = App.Path & "\Language\" & GetLanguageFileName

End Function

Public Function GetLanguageFileName() As String
    GetLanguageFileName = getConfig(CFG_SETTING_LANGUAGE)

    If (GetLanguageFileName = "") Then
        GetLanguageFileName = "english.ini"

    End If

End Function

Public Property Get GetFileNameByIndex(Index As Long) As String
    GetFileNameByIndex = colLanguages.Item("i" & Index)

End Property

Public Function EnumLanguageFile()

    Dim menuIndex As Long

    Dim tmpStr    As String

    tmpStr = Dir(App.Path & "\Language\*.ini")

    Do While tmpStr <> ""

        Dim strShowName As String

        strShowName = InI.INI_GetString(App.Path & "\Language\" & tmpStr, "Lang", "ShowName")
        colLanguages.Add tmpStr, "i" & menuIndex

        If (menuIndex > frmMenu.Language_Select.Count - 1) Then
            Load frmMenu.Language_Select(menuIndex)

        End If

        frmMenu.Language_Select.Item(menuIndex).Checked = False

        If (tmpStr = GetLanguageFileName) Then
            frmMenu.Language_Select.Item(menuIndex).Checked = True
            LanguageIndex = menuIndex

        End If

        frmMenu.Language_Select.Item(menuIndex).Caption = strShowName
        menuIndex = menuIndex + 1
        tmpStr = Dir()
    Loop

    If (menuIndex = 0) Then
        CreateLanguagePart frmMain
        CreateLanguagePart frmMenu
        CreateLanguagePart frmPaternAdd
        CreateLanguagePart frmSystemInfo

    End If

End Function

Public Sub SetLanguage(Index As Long)
    saveConfig "Language", GetFileNameByIndex(Index)
    LanguageIndex = Index
    
End Sub

Public Function ReApplyLanguage()
    ApplyLanguageToForm frmMain
    ApplyLanguageToForm frmMenu

End Function
