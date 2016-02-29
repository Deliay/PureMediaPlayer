VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   -360
   ClientLeft      =   150
   ClientTop       =   1125
   ClientWidth     =   840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   -360
   ScaleWidth      =   840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
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
    frmMain.isHide = Not frmMain.isHide
    
    If (frmMain.isHide) Then

    Else
        
    End If
    
End Sub

Private Sub mmStatus_SpeedDown_Click()
    frmMain.frmPlayer_KeyDown vbKeySubtract, 0

End Sub

Private Sub mmStatus_SpeedReset_Click()
    mdlGlobalPlayer.Rate = 100

End Sub

Private Sub mmStatus_SpeedUp_Click()
    frmMain.frmPlayer_KeyDown vbKeyAdd, 0
    
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

Public Sub Renderers_Click(Index As Integer)

    frmMenu.Renderers(val(getConfig("Renderer"))).Checked = False
    
    saveConfig "Renderer", CStr(Index)
    frmMenu.Renderers(Index).Checked = True
    
    GlobalRenderType = Index

End Sub
