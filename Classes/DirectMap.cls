VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colPairs As New Collection

Private strName  As String

Public Property Get Name() As String
    Name = strName

End Property

Public Function SetName(ByVal passstrName As String)
    strName = passstrName

End Function

Friend Function AddPair(pPair As DirectPair) As DirectPair
    colPairs.Add pPair, pPair.Key

End Function

Friend Function AddKeyValue(ByVal strKey As String, ByVal strValue As String)

    Dim tmp As New DirectPair

    tmp.InitPair strKey, strValue
    colPairs.Add tmp, strKey

End Function

Public Property Let Value(sKey As String, sVal As String)

    On Error GoTo NoExist

    colPairs.Item(sKey).Value = sVal

    Exit Property

NoExist:

    Dim newPair As New DirectPair

    newPair.InitPair sKey, sVal
    colPairs.Add newPair, sKey

    Resume

End Property

Public Function Exist(sKey As String)

    Dim objPair As DirectPair

    For Each objPair In colPairs

        If (objPair.Key = sKey) Then
            Exist = True

            Exit Function

        End If

    Next

    Exist = False

End Function

Public Property Get Value(sKey As String) As String
Attribute Value.VB_UserMemId = 0

    On Error GoTo NoExist

    Value = colPairs.Item(sKey).Value

    Exit Property

NoExist:

    Dim newPair As New DirectPair

    newPair.InitPair sKey, ""
    colPairs.Add newPair, sKey

    Resume

End Property

Public Property Get GetPair(ByVal sKey As String) As DirectPair
    Set GetPair = colPairs.Item(sKey)

End Property

Public Function Remove(ByVal strKey As String)
    colPairs.Remove strKey

End Function

'Public Function RemovePair2(ByVal objPair As DirectPair)
'    Dim objPairIter As DirectPair
'    For Each objPairIter In colPairs
'        If (ObjPtr(objPair) = ObjPtr(objPairIter)) Then
'            colPairs.Remove objPair.Key
'        End If
'    Next
'End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colPairs.[_NewEnum]

End Property

Public Property Get Count() As Long
    Count = colPairs.Count

End Property

