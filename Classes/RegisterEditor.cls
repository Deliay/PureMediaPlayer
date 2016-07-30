VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegisterEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function RegCloseKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegConnectRegistry _
                Lib "Advapi32.dll" _
                Alias "RegConnectRegistryA" (ByVal lpMachineName As String, _
                                             ByVal hKey As Long, _
                                             phkResult As Long) As Long

Private Declare Function RegCreateKey _
                Lib "Advapi32.dll" _
                Alias "RegCreateKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       phkResult As Long) As Long

Private Declare Function RegDeleteKey _
                Lib "Advapi32.dll" _
                Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue _
                Lib "Advapi32.dll" _
                Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                         ByVal lpValueName As String) As Long

Private Declare Function RegEnumKey _
                Lib "Advapi32.dll" _
                Alias "RegEnumKeyA" (ByVal hKey As Long, _
                                     ByVal dwIndex As Long, _
                                     ByVal lpName As String, _
                                     ByVal cbName As Long) As Long

Private Declare Function RegEnumValue _
                Lib "Advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       lpData As Byte, _
                                       lpcbData As Long) As Long

Private Declare Function RegFlushKey Lib "Advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKey _
                Lib "Advapi32.dll" _
                Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                     ByVal lpSubKey As String, _
                                     phkResult As Long) As Long

Private Declare Function RegOpenKeyEx _
                Lib "Advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
                Lib "Advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          lpData As Any, _
                                          lpcbData As Long) As Long ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function RegQueryValue _
                Lib "Advapi32.dll" _
                Alias "RegQueryValueA" (ByVal hKey As Long, _
                                        ByVal lpSubKey As String, _
                                        ByVal lpValue As String, _
                                        lpcbValue As Long) As Long

Private Declare Function RegReplaceKey _
                Lib "Advapi32.dll" _
                Alias "RegReplaceKeyA" (ByVal hKey As Long, _
                                        ByVal lpSubKey As String, _
                                        ByVal lpNewFile As String, _
                                        ByVal lpOldFile As String) As Long

Private Declare Function RegRestoreKey _
                Lib "Advapi32.dll" _
                Alias "RegRestoreKeyA" (ByVal hKey As Long, _
                                        ByVal lpFile As String, _
                                        ByVal dwFlags As Long) As Long

Private Declare Function RegSetValue _
                Lib "Advapi32.dll" _
                Alias "RegSetValueA" (ByVal hKey As Long, _
                                      ByVal lpSubKey As String, _
                                      ByVal dwType As Long, _
                                      ByVal lpData As String, _
                                      ByVal cbData As Long) As Long

Private Declare Function RegSetValueEx _
                Lib "Advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal Reserved As Long, _
                                        ByVal dwType As Long, _
                                        lpData As Any, _
                                        ByVal cbData As Long) As Long ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Private Declare Function RegUnLoadKey _
                Lib "Advapi32.dll" _
                Alias "RegUnLoadKeyA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String) As Long

Private Const ERROR_SUCCESS = 0&

Public Enum RegRootEnum

    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003

End Enum

Public Enum RegType

    REG_SZ = 1                                                                  ' Unicode nul terminated string
    REG_BINARY = 3                                                              ' Free form binary
    REG_DWORD = 4                                                               ' 32-bit number
    REG_EXPAND_SZ = 2                                                           ' Unicode nul terminated string
    REG_MULTI_SZ = 7                                                            ' Multiple Unicode strings

End Enum

Public Function GetString(hKey As RegRootEnum, strPath As String, strValue As String)

    Dim keyhand      As Long

    Dim lResult      As Long

    Dim strBuf       As String

    Dim lDataBufSize As Long

    Dim intZeroPos   As Integer

    Dim lValueType   As RegType                                                   'new add

    Dim iRet         As Long

    Dim iKeyHand     As Long

    Dim RegKeyRead   As String, ReturnStr As String

    iRet = RegOpenKey(hKey, strPath, iKeyHand)

    If iRet <> ERROR_SUCCESS Then
        RegKeyRead = vbNullString
        ReturnStr = vbNullString
        Exit Function

    End If

    lResult = RegQueryValueEx(iKeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Or lValueType = REG_EXPAND_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(iKeyHand, strValue, 0&, lValueType, ByVal strBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))

            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
                Else:       GetString = strBuf

            End If

        Else
            GetString = ""

        End If

    End If

End Function

Public Function SetString(hKey As RegRootEnum, _
                          strPath As String, _
                          strValue As String, _
                          strdata As String) As Boolean

    Dim keyhand As Long

    RegCreateKey hKey, strPath, keyhand

    If RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata)) <> ERROR_SUCCESS Then
        SetString = False
    Else
        SetString = True

    End If

    RegCloseKey keyhand

End Function

Function GetDword(ByVal hKey As RegRootEnum, _
                  ByVal strPath As String, _
                  ByVal strValueName As String) As Long

    Dim lResult      As Long

    Dim lValueType   As RegType

    Dim lBuf         As Long

    Dim lDataBufSize As Long

    Dim r            As Long

    Dim keyhand      As Long

    r = RegOpenKey(hKey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then
            GetDword = lBuf

        End If

        'Else
        ' Call errlog("GetDWORD-" & strPath, False)
    End If

    r = RegCloseKey(keyhand)

End Function

Function SetDword(ByVal hKey As RegRootEnum, _
                  ByVal strPath As String, _
                  ByVal strValueName As String, _
                  ByVal lData As Long) As Boolean

    Dim keyhand As Long

    RegCreateKey hKey, strPath, keyhand

    If RegSetValueEx(keyhand, strValueName, 0&, RegType.REG_DWORD, lData, 4) <> ERROR_SUCCESS Then
        SetDword = False
    Else
        SetDword = True

    End If
    
    RegCloseKey keyhand

End Function

Function GetBinary(ByVal hKey As RegRootEnum, _
                   ByVal strPath As String, _
                   ByVal strValueName As String) As Long

    Dim lResult      As Long

    Dim lValueType   As RegType

    Dim lBuf         As Long

    Dim lDataBufSize As Long

    Dim r            As Long

    Dim keyhand      As Long

    r = RegOpenKey(hKey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

    If lResult = ERROR_SUCCESS Then
        If lValueType = REG_BINARY Then
            GetBinary = lBuf

        End If

    End If

    r = RegCloseKey(keyhand)

End Function

Function SetBinary(ByVal hKey As RegRootEnum, _
                   ByVal strPath As String, _
                   ByVal strValueName As String, _
                   ByVal lData As Long, _
                   ByVal BitNumber As Long) As Boolean

    Dim keyhand As Long

    RegCreateKey hKey, strPath, keyhand

    If RegSetValueEx(keyhand, strValueName, 0&, REG_BINARY, lData, BitNumber) <> ERROR_SUCCESS Then
        SetBinary = False
    Else
        SetBinary = True

    End If

    RegCloseKey keyhand

End Function

Public Function DeleteValue(ByVal hKey As RegRootEnum, _
                            ByVal strPath As String, _
                            ByVal strValue As String) As Boolean

    Dim keyhand As Long

    RegOpenKey hKey, strPath, keyhand

    If RegDeleteValue(keyhand, strValue) <> ERROR_SUCCESS Then
        DeleteValue = False
    Else
        DeleteValue = True

    End If
    
    RegCloseKey keyhand

End Function

Public Function CreateKey(ByVal hKey As RegRootEnum, ByVal strKey As String) As Boolean

    Dim keyhand&

    If RegCreateKey(hKey, strKey, keyhand) <> ERROR_SUCCESS Then
        CreateKey = False
    Else
        CreateKey = True

    End If
    
    RegCloseKey keyhand&

End Function

Public Function ItemExits(hKey As RegRootEnum, Key As String) As Boolean
    
    Dim Ret As Long, xKey As Long

    If RegOpenKey(hKey, Key, xKey) = 0& Then
        ItemExits = True
    Else
        ItemExits = False

    End If

End Function

Public Function DelKey(hKey As RegRootEnum, Key As String) As Boolean

    Dim sKey As Long, Ret As Long, Name As String, Idx As Long
    
    If RegOpenKey(hKey, Key, sKey) <> ERROR_SUCCESS Then
        DelKey = True
    Else
        DelKey = False

    End If
    
    Idx = 0
    Name = String(256, Chr(0))
    Do
        Ret = RegEnumKey(sKey, Idx, Name, Len(Name))

        If Ret = 0 Then
            DeleteValue hKey, Key, Name

        End If

    Loop Until Ret <> 0
    
    If RegDeleteKey(hKey, Key) <> ERROR_SUCCESS Then
        DelKey = False
    Else
        DelKey = True

    End If
    
End Function
