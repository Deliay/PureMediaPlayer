VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Pure Media Player"
   ClientHeight    =   3960
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   12930
   Icon            =   "pmdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer tmrUpdateTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   3000
   End
   Begin VB.PictureBox pbTimeBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   10
      ScaleMode       =   0  'User
      ScaleWidth      =   862
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3510
      Width           =   12930
      Begin VB.PictureBox pbTimeBlock 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         ScaleHeight     =   9
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   25
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3645
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   882
            Text            =   "Wait"
            TextSave        =   "Wait"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   353
            MinWidth        =   353
            Text            =   "|"
            TextSave        =   "|"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1402
            MinWidth        =   882
            Text            =   "Stoped."
            TextSave        =   "Stoped."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   353
            MinWidth        =   353
            Text            =   "|"
            TextSave        =   "|"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2408
            MinWidth        =   882
            Text            =   "0% (0:00/0:00)"
            TextSave        =   "0% (0:00/0:00)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   353
            MinWidth        =   353
            Text            =   "|"
            TextSave        =   "|"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2646
            MinWidth        =   882
            Text            =   "No File Open....."
            TextSave        =   "No File Open....."
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   12656
            MinWidth        =   882
            TextSave        =   "2016/2/23"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1032
            MinWidth        =   882
            TextSave        =   "14:46"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
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
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
    mdlGlobalPlayer.SwitchFullScreen True, False
    
End Sub

Private Sub MDIForm_Load()
    Me.Show
    Load frmPlaylist
    Load frmPaternAdd
    
    If (Dir(App.Path & "\language.ini") = "") Then
        CreateLanguagePart Me
        CreateLanguagePart frmPlaylist
        CreateLanguagePart frmPaternAdd
    Else
        ApplyLanguageToForm Me
        ApplyLanguageToForm frmPlaylist
        ApplyLanguageToForm frmPaternAdd

    End If

    frmPlaylist.Hide
    frmPaternAdd.Hide
    'mdlGlobalPlayer.LoadMediaFile test
    frmPlayer.Show
    frmPlayer.WindowState = 2
    frmPlayer.Caption = StaticString(PLAYER_STATUS_IDLE)
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    UpdateStatus StaticString(FILE_STATUS_NOFILE), StatusBarEnum.FileName
    Me.Height = 8300
    Me.Width = 12800

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ExitProgram
    
End Sub

Private Sub MDIForm_Resize()
    mdlGlobalPlayer.Width = frmPlayer.Width / Screen.TwipsPerPixelX
    mdlGlobalPlayer.Height = frmPlayer.Height / Screen.TwipsPerPixelY
    ResizePlayWindow
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ExitProgram
    
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
        frmPlaylist.isHide = True
        frmPlaylist.Show vbModeless, frmMain
        frmPlaylist.AutoPatern
        frmPlaylist.AutoPatern
        frmPlaylist.File_PaternFind.Enabled = True
        RenderMediaFile
        Me.Width = Me.Width + 10

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

Private Sub CalcPercent(X As Single)
    Precent = (X / pbTimeBar.Width) * 100
    tmrUpdateTime_Timer
    
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
    frmPlaylist.isHide = Not frmPlaylist.isHide
    
    If (frmPlaylist.isHide) Then
        frmPlaylist.Show vbModeless, Me
    Else
        frmPlaylist.Hide
        
    End If
    
End Sub

Private Sub mmStatus_SpeedDown_Click()
    frmPlayer.Form_KeyDown vbKeySubtract, 0
End Sub

Private Sub mmStatus_SpeedReset_Click()
    mdlGlobalPlayer.Rate = 100
End Sub

Private Sub mmStatus_SpeedUp_Click()
    frmPlayer.Form_KeyDown vbKeyAdd, 0
    
End Sub

Private Sub mmStatus_Stop_Click()
    StopPlay
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    
End Sub

Private Sub pbTimeBar_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    CalcPercent X * Screen.TwipsPerPixelX
    frmPlayer.SetFocus
    tmrUpdateTime.Enabled = True
    
End Sub

Private Sub pbTimeBlock_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    CalcPercent X * Screen.TwipsPerPixelX
    frmPlayer.SetFocus
    tmrUpdateTime.Enabled = True
    
End Sub

Private Sub tmrUpdateTime_Timer()
    Sleep 1

    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    UpdateStatus mdlGlobalPlayer.Volume & "%, " & Round(mdlGlobalPlayer.Rate / 100, 2) & "x" & ", " & Format(Round(Precent, 2), "##.#0") & "% (" & FormatedCurrentTime & "/" & FormatedDuration & ")", PlayTime
    pbTimeBlock.Width = Precent / 100 * (pbTimeBar.Width / Screen.TwipsPerPixelX)
    
    If (Duration < CurrentTime + 1) Then
        tmrUpdateTime.Enabled = False
        
        If (mdlPlaylist.PlaylistPlayNext) Then
            CurrentTime = 0
            pbTimeBlock.Width = 0
        Else
            
            If (mmStatus_Loop.Checked) Then
                CurrentTime = 0
                pbTimeBlock.Width = 0
                
            End If
            
        End If
        
    End If

    'Set Me.Picture = mdlGlobalPlayer.NowFrame
End Sub
