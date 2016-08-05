VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private colList As New Collection

Private strName As String

Public Property Get Name() As String
    Name = strName

End Property

Public Function SetName(ByVal passstrName As String)
    strName = passstrName

End Function

Friend Function AddItem(ByVal strItem As String)
    colList.Add strItem
    
End Function

Public Property Get Item(ByVal Index As Long) As String
Attribute Item.VB_UserMemId = 0
    Item = colList(Index)

End Property

Friend Function Remove(ByVal Index As Long)
    colList.Remove Index

End Function

Public Sub Clear()
    Set colList = New Collection
    colList.Add "@"
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colList.[_NewEnum]
    
End Property

Public Property Get Count() As Long
    Count = colList.Count

End Property

Public Function TakeTo(List As Control)

    Dim i As Variant

    For Each i In colList

        List.AddItem i
    Next

End Function

Public Function Exist(Item As String) As Boolean

    Dim i As Variant

    For Each i In colList

        If (i = Item) Then
            Exist = True
            Exit Function
        End If
    Next
    
    Exist = False

End Function

Public Function TakeFrom(List As Control)

    Dim lngCount As Long

    Dim lngIter  As Long

    Set colList = New Collection
    lngCount = List.ListCount - 1

    For lngIter = 0 To lngCount
        colList.Add List.List(lngIter)
    Next

End Function
