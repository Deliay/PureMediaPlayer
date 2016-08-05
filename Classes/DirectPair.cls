VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private strKey   As String

Private strValue As String

Public Property Get Key() As String
    Key = strKey

End Property

Public Property Get Value() As String
    Value = strValue

End Property

Public Property Let Key(v As String)
    strKey = v

End Property

Public Property Let Value(v As String)
    strValue = v

End Property

Friend Sub InitPair(ByVal sKey As String, ByVal sValue As String)
    strKey = sKey
    strValue = sValue
    
End Sub

Friend Sub InitPairArr(ByRef splitedArr)
    strKey = splitedArr(0)
    strValue = splitedArr(1)
    
End Sub

