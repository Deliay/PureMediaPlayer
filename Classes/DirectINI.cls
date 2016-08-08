VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strFilePath    As String

Private lngFileNum     As Long

Private strFileContent As String

Private idBind         As Object

Private colMaps        As New Collection

Private colLists       As New Collection

Private colPairs       As New Collection

Public Function CreateMap(ByVal Name As String) As DirectMap

    Dim objMap As New DirectMap

    objMap.SetName Name
    colMaps.Add objMap, Name
    Set CreateMap = objMap
    CreateMap.AddPair CreateKeyValuePair("@", "@")

End Function

Public Function CreateList(ByVal Name As String) As DirectList

    Dim objList As New DirectList

    objList.SetName Name
    colLists.Add objList, Name
    Set CreateList = objList
    CreateList.AddItem "@"

End Function

Public Function CreateKeyValuePair(ByVal Key As String, _
                                   ByVal Value As String) As DirectPair

    Dim objPair As New DirectPair

    objPair.InitPair Key, Value
    Set CreateKeyValuePair = objPair

End Function

Public Function CreatePair(ByVal Key As String, ByVal Value As String) As DirectPair

    Dim objPair As New DirectPair

    objPair.InitPair Key, Value
    colPairs.Add objPair, Key
    Set CreatePair = objPair

End Function

Public Function SaveToOpendFile()

    If (lngFileNum > 0 And Len(strFilePath) > 0) Then
        Open File For Output As #lngFileNum
        Print #1, SaveToString()
        Close #lngFileNum

    End If

End Function

Public Function SaveToString() As String

    Dim k As DirectPair, map As DirectMap, List As DirectList

    If (idBind Is Nothing) Then

        For Each k In colPairs

            SaveToString = SaveToString & k.Key & "=" & k.Value & vbCrLf
        Next
    Else

        For Each k In colPairs

            SaveToString = SaveToString & k.Key & "=" & CallByName(idBind, k.Key, VbGet) & vbCrLf
        Next

    End If
    
    Dim objPairs As DirectPair

    For Each map In colMaps

        SaveToString = SaveToString & "[" & map.Name & "]" & vbCrLf

        For Each objPairs In map

            SaveToString = SaveToString & objPairs.Key & "=" & objPairs.Value & vbCrLf
        Next
    Next
    
    Dim strListValue As Variant

    For Each List In colLists

        SaveToString = SaveToString & "[" & List.Name & "]" & vbCrLf

        For Each strListValue In List

            SaveToString = SaveToString & "=" & strListValue & vbCrLf
        Next
    Next
    
End Function

Private Function OpenFile(ByVal strPath As String)
    lngFileNum = FreeFile()
    Open strPath For Binary As #lngFileNum
    strFilePath = strPath

End Function

Public Property Get File() As String
    File = strFilePath

End Property

Public Function ReadFormFile(ByVal strPath As String) As Boolean

    Dim strFileContent As String

    OpenFile strPath
    strFileContent = Space(LOF(lngFileNum))
    Get #lngFileNum, , strFileContent
    ReadFormString strFileContent
    Close #1

End Function

Public Function ReadFormString(ByVal strContent As String) As Boolean
    
    Dim astrLines() As String

    astrLines = Split(strContent, vbCrLf)
    
    Dim lngLines As Long, lngSizeLines As Long

    lngSizeLines = UBound(astrLines)
    
    Dim bFlagIsInMap   As Boolean, bFlagIsInList As Boolean

    Dim strName        As String

    Dim dCurrentMap    As DirectMap

    Dim dCurrentList   As DirectList

    Dim strCurrentLine As String

    For lngLines = 0 To lngSizeLines
        strCurrentLine = Trim$(astrLines(lngLines))

        If (InStr(1, strCurrentLine, Chr(0)) > 0) Then strCurrentLine = Left$(strCurrentLine, InStr(1, strCurrentLine, Chr(0)) - 1)
        If (strCurrentLine = "") Then GoTo NextLine
        If (Left$(strCurrentLine, 1) = "[") Then
            If (Left$(Trim$(astrLines(lngLines + 1)), 1) = "=") Then
                bFlagIsInList = True
                bFlagIsInMap = False
                Set dCurrentList = New DirectList
                strName = Mid(strCurrentLine, 2, Len(strCurrentLine) - 2)
                colLists.Add dCurrentList, strName
                dCurrentList.SetName strName
            ElseIf lngLines + 1 > UBound(astrLines) Then
                Exit For
            ElseIf Len(Trim$(astrLines(lngLines + 1))) = 0 Then
                GoTo NextLine
            Else
                bFlagIsInMap = True
                bFlagIsInList = False
                Set dCurrentMap = New DirectMap
                strName = Mid(strCurrentLine, 2, Len(strCurrentLine) - 2)
                colMaps.Add dCurrentMap, strName
                dCurrentMap.SetName strName

            End If

        Else

            If (bFlagIsInList) Then
                dCurrentList.AddItem Mid$(strCurrentLine, 2)
            Else

                Dim dPair As DirectPair

                Set dPair = New DirectPair
                dPair.InitPairArr Split(strCurrentLine, "=", 2)

                If (bFlagIsInMap) Then
                    dCurrentMap.AddPair dPair
                Else
                    colPairs.Add dPair, dPair.Key

                End If

            End If

        End If

NextLine:
    Next
    
End Function

Public Function GetValue(ByVal strKey As String, _
                         Optional ByVal strSub As String = ".") As String

    If (strSub = ".") Then
        GetValue = colPairs(strKey).Value
    Else
        GetValue = colMaps(strSub)(strKey)

    End If
    
End Function

Public Function SetValue(ByVal strKey As String, _
                         ByVal strValue As String, _
                         Optional ByVal strSub As String = ".")

    If (idBind Is Nothing) Then
        If (strSub = ".") Then
            colPairs(strKey).Value = strValue
        Else
            colMaps(strSub)(strKey) = strValue

        End If

    Else
        
        If (strSub = ".") Then
            CallByName idBind, strKey, VbLet, strValue
        Else
            colMaps(strSub)(strKey) = strValue

        End If

    End If

End Function

Public Function GetMap(ByVal strMapName As String) As DirectMap
    Set GetMap = colMaps(strMapName)
    
End Function

Public Function GetList(ByVal strListName As String) As DirectList
    Set GetList = colLists(strListName)
    
End Function

Public Function Bind(ByVal idTarget As Object) As Boolean
    Set idBind = idTarget

    Dim k As DirectPair, map As DirectMap, List As DirectList

    '    Dim m_rrid As UUID, m_lcid As Long, m_dispid As Long
    '    idTarget.GetIDsOfNames m_rrid, "Key", 1, m_lcid, m_dispid
    '    Dim m_Param As olelib.DISPPARAMS, m_e As olelib.EXCEPINFO
    '    Dim m_Str As String
    '    m_Str = Space(256)
    '    m_Param
    '
    '    idTarget.Invoke m_dispid, m_rrid, m_lcid, DISPATCH_PROPERTYGET, m_Param, StrPtr(m_Str), m_e, ByVal 0
    
    idBind.InitBind Me
    
    For Each k In colPairs

        CallByName idBind, k.Key, VbLet, k.Value
    Next
    
    For Each map In colMaps

        CallByName idBind, map.Name, VbSet, map
    Next
    
    For Each List In colLists

        CallByName idBind, List.Name, VbSet, List
    Next

End Function
