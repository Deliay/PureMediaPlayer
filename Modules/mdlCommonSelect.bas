Attribute VB_Name = "cdlg"
Option Explicit

Public Const CFG_HISTORY_LAST_SAVE_PATH As String = "LastSavePath"

Public Const CFG_HISTORY_LAST_OPEN_PATH As String = "LastOpenPath"

Private Declare Function GetSaveFileName _
                Lib "comdlg32.dll" _
                Alias "GetSaveFileNameW" (ByVal pOpenfilename As Long) As Long

Private Declare Function GetOpenFileName _
                Lib "comdlg32.dll" _
                Alias "GetOpenFileNameW" (ByVal pOpenfilename As Long) As Long

Private Type OPENFILENAME

    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As Long
    lpstrCustomFilter As Long
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As Long
    nMaxFileTitle As Long
    lpstrInitialDir As Long
    lpstrTitle As Long
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long

End Type

Dim ofn              As OPENFILENAME

Dim rtn              As String

Public FileName      As String

Public FileNameWiden As String

Private Declare Function GetShortPathName _
                Lib "kernel32" _
                Alias "GetShortPathNameW" (ByVal lpszLongPath As Long, _
                                           ByVal lpszShortPath As Long, _
                                           ByVal cchBuffer As Long) As Long

Public Function ShowOpen(Optional Filters As String = "任意文件 (*.*)" & vbNullChar & "*.*")
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = frmMain.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = StrPtr(Filters)
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = StrPtr(Space(254))
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = StrPtr(getConfig(CFG_HISTORY_LAST_OPEN_PATH, History))
    ofn.lpstrTitle = StrPtr("Please Select File")
    ofn.flags = 6148
    rtn = GetOpenFileName(VarPtr(ofn))
    
    If rtn < 1 Then Exit Function
    FileName = Space$(254)
    rtn = GetShortPathName(StrPtr(ofn.lpstrFile), StrPtr(FileName), 254)

    If (Not (InStr(1, FileName, Chr(0)) = 0)) Then
        FileName = Mid(FileName, 1, InStr(1, FileName, Chr(0)) - 1)

    End If
 
    saveConfig "LastOpenDir", DirGet(FileName), History
    
End Function

Public Function ShowSave(Optional Filters As String = "任意文件 (*.*)" & vbNullChar & "*.*")
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = frmMain.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = StrPtr(Filters)
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = StrPtr(Space(254))
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = StrPtr(getConfig(CFG_HISTORY_LAST_SAVE_PATH, History))
    ofn.lpstrTitle = StrPtr("Please Select File")
    ofn.flags = 6148
    rtn = GetSaveFileName(VarPtr(ofn))
    
    If rtn < 1 Then Exit Function
    FileName = Space$(254)
    rtn = GetShortPathName(StrPtr(ofn.lpstrFile), StrPtr(FileName), 254)
    If (Left$(FileName, 1) = Space$(1)) Then FileName = ofn.lpstrFile
    If (Not (InStr(1, FileName, Chr(0)) = 0)) Then
        FileName = Mid(FileName, 1, InStr(1, FileName, Chr(0)) - 1)

    End If
 
    saveConfig "LastSaveDir", DirGet(FileName), History
    
End Function

Public Function DirGet(ByVal strFilePath As String) As String
    
    If (strFilePath = "") Then Exit Function
    DirGet = Mid$(strFilePath, 1, InStrRev(strFilePath, "\") - 1)
    
End Function

Public Function NameGet(ByVal strFilePath As String) As String
    
    If (strFilePath = "") Then Exit Function
    NameGet = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    
End Function

Private Function ConvertFileName(sToConvert As String) As String
   
End Function

