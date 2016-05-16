VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "PureMediaPlayerMenuHost"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   2520
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   62
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   Visible         =   0   'False
   Begin VB.Menu MenuMain 
      Caption         =   "Menu"
      Begin VB.Menu mmFile 
         Caption         =   "File"
         Begin VB.Menu mmFile_Open 
            Caption         =   "Open"
            Shortcut        =   ^D
         End
         Begin VB.Menu mmFile_Close 
            Caption         =   "Close"
            Shortcut        =   ^C
         End
         Begin VB.Menu mmFile_Spec1 
            Caption         =   "-"
         End
         Begin VB.Menu mmFile_Option 
            Caption         =   "Option"
            Enabled         =   0   'False
         End
         Begin VB.Menu mmFile_Spec3 
            Caption         =   "-"
         End
         Begin VB.Menu mmFile_Exit 
            Caption         =   "Exit"
            Shortcut        =   ^X
         End
      End
      Begin VB.Menu mmStatus 
         Caption         =   "Status"
         Begin VB.Menu mmStatus_Play 
            Caption         =   "Play"
         End
         Begin VB.Menu mmStatus_Stop 
            Caption         =   "Stop"
         End
         Begin VB.Menu mmStatus_Pause 
            Caption         =   "Pause"
         End
         Begin VB.Menu mmStatus_Spec1 
            Caption         =   "-"
         End
         Begin VB.Menu mmStatus_Random 
            Caption         =   "Random"
            Enabled         =   0   'False
         End
         Begin VB.Menu mmStatus_ShowPlaylist 
            Caption         =   "Playlist"
            Shortcut        =   ^F
         End
         Begin VB.Menu mmStatus_Loop 
            Caption         =   "Loop"
            Checked         =   -1  'True
         End
         Begin VB.Menu mmStatus_Spec2 
            Caption         =   "-"
         End
         Begin VB.Menu mmStatus_SpeedUp 
            Caption         =   "Speed+0.1x"
         End
         Begin VB.Menu mmStatus_SpeedDown 
            Caption         =   "Speed-0.1x"
         End
         Begin VB.Menu mmStatus_SpeedReset 
            Caption         =   "Reset Speed 1.0x"
         End
      End
      Begin VB.Menu Renderer 
         Caption         =   "Renderer"
         Begin VB.Menu Renderers 
            Caption         =   "Video Renderer"
            Index           =   0
         End
         Begin VB.Menu Renderers 
            Caption         =   "VMR7"
            Index           =   1
         End
         Begin VB.Menu Renderers 
            Caption         =   "VRM9(Windowless)"
            Index           =   2
         End
         Begin VB.Menu Renderers 
            Caption         =   "EVR(CP)"
            Index           =   3
         End
         Begin VB.Menu Renderers 
            Caption         =   "MADVR"
            Index           =   4
         End
      End
      Begin VB.Menu mmInfo 
         Caption         =   "Info"
         Begin VB.Menu mmInfo_Media 
            Caption         =   "Media"
         End
         Begin VB.Menu mmInfo_System 
            Caption         =   "System"
         End
         Begin VB.Menu mmInfo_Software 
            Caption         =   "Software"
         End
      End
      Begin VB.Menu Language 
         Caption         =   "Language"
         Begin VB.Menu Language_Select 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu Propertys 
         Caption         =   "Property"
         Begin VB.Menu Propertys_Video 
            Caption         =   "Video"
         End
         Begin VB.Menu Propertys_Audio 
            Caption         =   "Audio"
         End
         Begin VB.Menu Propertys_Splitter 
            Caption         =   "Splitter"
         End
         Begin VB.Menu Propertys_Renderer 
            Caption         =   "Renderer"
         End
         Begin VB.Menu Propertys_Split1 
            Caption         =   "-"
         End
         Begin VB.Menu Propertys_Subtitle 
            Caption         =   "Subtitle"
         End
      End
      Begin VB.Menu mmHelp 
         Caption         =   "Help"
         Begin VB.Menu mmHelp_About 
            Caption         =   "About RnW"
         End
         Begin VB.Menu mmHelp_Web 
            Caption         =   "Website"
         End
         Begin VB.Menu mmHelp_Spec 
            Caption         =   "-"
         End
         Begin VB.Menu mmHelp_Help 
            Caption         =   "Help"
         End
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private Sub Language_Select_Click(Index As Integer)

    If (mdlLanguageApplyer.GetLanguageName = Language_Select(Index).Caption) Then
        'same item reclicked
        Exit Sub
    Else
        'else item click

        Language_Select(LanguageIndex).Checked = False
        Language_Select(Index).Checked = True
        mdlLanguageApplyer.SetLanguage CLng(Index)
        mdlLanguageApplyer.ReApplyLanguage

    End If
    
End Sub

Private Sub mmHelp_About_Click()
    MsgBox "Remilia(Net) Workstation(admin@remiliascarlet.com)"

End Sub

Private Sub mmHelp_Help_Click()
    MsgBox "Free to use"

End Sub

Private Sub mmHelp_Web_Click()
    ShellExecute 0, "open", "http://www.remiliascarlet.com", vbNullString, vbNullString, 0
    
End Sub

Private Sub mmStatus_Pause_Click()
    Pause
    UpdateStatus StaticString(PLAY_STATUS_PAUSED), PlayBack
    
End Sub

Private Sub mmStatus_Play_Click()
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    Play
    UpdateStatus StaticString(PLAY_STATUS_PLAYING), PlayBack
    
End Sub

Private Sub mmStatus_ShowPlaylist_Click()

    If (boolPlaylistStatus = True) Then
        PlaylistHide
    Else
        PlaylistShow

        'RefreshUI
    End If
    
End Sub

Private Sub mmStatus_SpeedDown_Click()
    frmMain.Form_KeyDown vbKeySubtract, 0

End Sub

Private Sub mmStatus_SpeedReset_Click()
    mdlGlobalPlayer.Rate = 100

End Sub

Private Sub mmStatus_SpeedUp_Click()
    frmMain.Form_KeyDown vbKeyAdd, 0
    
End Sub

Private Sub mmStatus_Stop_Click()
    StopPlay
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    
End Sub

Private Sub mmFile_Close_Click()
    StopPlay
    CloseFile
    
End Sub

Private Sub mmFile_Exit_Click()
    ExitProgram
    
End Sub

Private Sub mmFile_Open_Click()
    cdlg.FileName = ""
    cdlg.ShowOpen
    
    If (Len(cdlg.FileName) > 0 And Dir(cdlg.FileNameWiden) <> "") Then
        mdlGlobalPlayer.CloseFile
        mdlGlobalPlayer.File = cdlg.FileName
        frmMain.isHide = True
        frmMain.AutoPatern
        RenderMediaFile

    End If
    
End Sub

Private Sub mmInfo_Media_Click()
    frmMediaInfo.Show
    
End Sub

Private Sub mmInfo_Software_Click()
    MsgBox "Pure Media Player (PMP) ver " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub mmInfo_System_Click()
    frmSystemInfo.Show
    
End Sub

Private Sub Propertys_Audio_Click()
    mdlFilterBuilder.ShowAudioDecoderConfig

End Sub

Private Sub Propertys_Renderer_Click()
    mdlFilterBuilder.ShowRendererConfig

End Sub

Private Sub Propertys_Splitter_Click()
    mdlFilterBuilder.ShowSpliterConfig

End Sub

Private Sub Propertys_Subtitle_Click()
    mdlFilterBuilder.ShowSubtitleConfig

End Sub

Private Sub Propertys_Video_Click()
    mdlFilterBuilder.ShowVideoDecoderConfig

End Sub

Public Sub Renderers_Click(Index As Integer)

    frmMenu.Renderers(val(getConfig("Renderer"))).Checked = False
    
    saveConfig "Renderer", CStr(Index)
    frmMenu.Renderers(Index).Checked = True

End Sub
