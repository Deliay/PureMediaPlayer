VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LastPlayPos  As DirectMap

Public LastPlayList As DirectList

Public FileDuration As DirectMap

Public SubtitleBind As DirectMap

Public Language     As String

Public Renderer     As String

Public LastHwnd     As String

Public LastOpenDir  As String

Public LastSaveDir  As String

Public Sub InitBind(binder As DirectINI)

    On Error Resume Next

    With binder
        .CreateMap "LastPlayPos"
        .CreateMap "FileDuration"
        .CreateMap "SubtitleBind"
        .CreateList "LastPlayList"
        .CreatePair "Language", "chinese.ini"
        .CreatePair "Renderer", "4"
        .CreatePair "LastHwnd", "0"
        .CreatePair "LastOpenDir", "."
        .CreatePair "LastSaveDir", "."

    End With

    On Error GoTo 0

End Sub
