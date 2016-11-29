Attribute VB_Name = "mdlConfig"
Option Explicit

Public GlobalConfig  As New Config

Private objINIReader As New DirectINI

Public Sub InitConfigFiles()
    'Environ$("LocalAppData")
    'TODO: Save config in AppData/Local/PureMediaPlayer
    objINIReader.ReadFormFile App.Path & "\Config.ini"
    objINIReader.Bind GlobalConfig

End Sub

Public Sub SaveConfig()
    objINIReader.SaveToOpendFile

End Sub
