Attribute VB_Name = "cdlg"
Option Explicit

Public Const CFG_HISTORY_LAST_SAVE_PATH As String = "LastSavePath"

Public Const CFG_HISTORY_LAST_OPEN_PATH As String = "LastOpenPath"

Private Declare Function GetSaveFileName _
                Lib "comdlg32.dll" _
                Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetOpenFileName _
                Lib "comdlg32.dll" _
                Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME

    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String

End Type

Dim ofn         As OPENFILENAME

Dim rtn         As String

Public FileName As String

Public Function ShowOpen(Optional Filters As String = "任意文件 (*.*)" & vbNullChar & "*.*")
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = frmMain.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = Filters
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = getConfig(CFG_HISTORY_LAST_OPEN_PATH, History)
    ofn.lpstrTitle = "Please Select File"
    ofn.flags = 6148
    rtn = GetOpenFileName(ofn)
    
    If rtn < 1 Then Exit Function
    If (Not (InStr(1, ofn.lpstrFile, Chr(0)) = 0)) Then
        ofn.lpstrFile = Mid(ofn.lpstrFile, 1, InStr(1, ofn.lpstrFile, Chr(0)) - 1)
        
    End If
    
    FileName = ofn.lpstrFile
    saveConfig "LastOpenDir", DirGet(FileName), History
    
End Function

Public Function ShowSave(Optional Filters As String = "任意文件 (*.*)" & vbNullChar & "*.*")
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = frmMain.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = Filters
    ofn.lpstrFile = Space(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = getConfig(CFG_HISTORY_LAST_SAVE_PATH, History)
    ofn.lpstrTitle = "Please Select File"
    ofn.flags = 6148
    rtn = GetSaveFileName(ofn)
    
    If rtn < 1 Then Exit Function
    If (Not (InStr(1, ofn.lpstrFile, Chr(0)) = 0)) Then
        ofn.lpstrFile = Mid(ofn.lpstrFile, 1, InStr(1, ofn.lpstrFile, Chr(0)) - 1)
        
    End If
    
    FileName = ofn.lpstrFile
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
