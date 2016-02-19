Attribute VB_Name = "InI"
Option Explicit
                        
Private Declare Function WritePrivateProfileString _
                Lib "kernel32" _
                Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                    ByVal lpKeyName As Any, _
                                                    ByVal lpString As Any, _
                                                    ByVal lpFileName As String) As Long
        
'Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
        
Private Declare Function GetPrivateProfileString _
                Lib "kernel32" _
                Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                  ByVal lpKeyName As Any, _
                                                  ByVal lpDefault As String, _
                                                  ByVal lpReturnedString As String, _
                                                  ByVal nSize As Long, _
                                                  ByVal lpFileName As String) As Long

Public Function INI_WriteString(FilePath As String, Section, KeyName, KeyString)
    WritePrivateProfileString CStr(Section), CStr(KeyName), CStr(KeyString), FilePath
    
End Function

Public Function INI_GetString(FilePath As String, _
                              Section As String, _
                              KeyName As String) As String
    
    Dim xSize As Long, kTemp As String, s As Integer

    xSize = 255
    kTemp = String(xSize, 0)
    GetPrivateProfileString Section, KeyName, "", kTemp, xSize, FilePath
    s = InStr(kTemp, Chr(0))

    If s > 0 Then
        kTemp = Left(kTemp, s - 1)
        INI_GetString = kTemp

    End If

End Function

