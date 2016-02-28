VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "Pure Media Player"
   ClientHeight    =   6705
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   10455
   Icon            =   "pmdiMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox bbMenuBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   480
   End
   Begin VB.Timer tmrUpdateTime 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   5760
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
      ScaleWidth      =   697
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6255
      Width           =   10455
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
      Top             =   6390
      Width           =   10455
      _ExtentX        =   18441
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
            Object.Width           =   8290
            MinWidth        =   882
            TextSave        =   "2016/2/28"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1032
            MinWidth        =   882
            TextSave        =   "13:27"
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
   Begin VB.PictureBox frmPlayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   10455
      Begin VB.PictureBox bbPlaylist 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   10080
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   24
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2760
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bbMenuBar_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    'DragWindow Me.hwnd
    Me.PopupMenu frmMenu.MenuMain

End Sub

Private Sub bbMenuBar_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If (bbMenuBar.BackColor <> RGB(48, 48, 48)) Then
        bbMenuBar.BackColor = RGB(48, 48, 48)
        SwitchUI True

    End If

End Sub

Private Sub bbPlaylist_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (bbPlaylist.BackColor <> RGB(48, 48, 48)) Then
        bbPlaylist.BackColor = RGB(48, 48, 48)
        SwitchUI True

    End If

End Sub

Private Sub Form_Activate()
    mdlGlobalPlayer.SwitchFullScreen True, False
    
End Sub

Private Sub Form_Load()
    Me.Show
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    UpdateStatus StaticString(FILE_STATUS_NOFILE), StatusBarEnum.FileName
    SwitchUI True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeWindow Me.hWnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bbMenuBar.BackColor = vbBlack

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ExitProgram
    
End Sub

Private Sub Form_Resize()
    Sleep 1
    ReCalcPlayWindow
    RefreshUI
    DoEvents
End Sub

Public Sub ReCalcPlayWindow()
    frmPlayer.width = (Me.width / Screen.TwipsPerPixelX)
    frmPlayer.height = (Me.height / Screen.TwipsPerPixelY)
    mdlGlobalPlayer.width = frmPlayer.width
    mdlGlobalPlayer.height = frmPlayer.height - mdlToolBarAlphaer.UIHeightButtom - 31
    bbPlaylist.Left = frmPlayer.width - 32
    bbPlaylist.Top = frmPlayer.height / 2 - bbPlaylist.height
    ResizePlayWindow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExitProgram
    
End Sub

Private Sub frmPlayer_Click()
    SwitchUI

End Sub

Private Sub frmPlayer_DblClick()
    SwitchFullScreen
    SwitchUI True

End Sub

Public Sub frmPlayer_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = vbKeySpace) Then SwitchPlayStauts
    If (KeyCode = vbKeyEscape) Then SwitchFullScreen True
    
    Dim flag As Long
    
    flag = 0
    
    If (mdlGlobalPlayer.Loaded) Then
        If (KeyCode = vbKeyLeft) Then flag = -1
        If (KeyCode = vbKeyRight) Then flag = 1
        If (flag <> 0) Then
            mdlGlobalPlayer.CurrentTime = mdlGlobalPlayer.CurrentTime + 5 * flag
            Exit Sub

        End If

        If (KeyCode = vbKeyUp) Then flag = 1
        If (KeyCode = vbKeyDown) Then flag = -1
        If (flag <> 0) Then
            mdlGlobalPlayer.Volume = mdlGlobalPlayer.Volume + flag
            Exit Sub

        End If

        If (KeyCode = vbKeyAdd) Then flag = 1
        If (KeyCode = vbKeySubtract) Then flag = -1
        If (flag <> 0) Then
            mdlGlobalPlayer.Rate = mdlGlobalPlayer.Rate + flag * 10
            Exit Sub

        End If

    End If
    
End Sub

Private Sub frmPlayer_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    If (Button = vbRightButton) Then Me.PopupMenu frmMenu.mmStatus

End Sub

Private Sub CalcPercent(X As Single)
    Precent = (X / pbTimeBar.width) * 100
    tmrUpdateTime_Timer
    
End Sub

Private Sub frmPlayer_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If (bbMenuBar.BackColor <> 0) Then
        bbMenuBar.BackColor = 0
        mdlToolBarAlphaer.apMenuButton.RefreshHW 32, 32

    End If
    If (bbPlaylist.BackColor <> 0) Then
        bbPlaylist.BackColor = 0
        mdlToolBarAlphaer.apPlaylistHint.RefreshHW 24, 48

    End If
    If (X < 32 And Y < 32) Or (X > bbPlaylist.Left And Y > bbPlaylist.Top And Y < bbPlaylist.Top + bbPlaylist.height) Then
        RefreshUI
    End If

End Sub

Private Sub pbTimeBar_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    CalcPercent X
    frmPlayer.SetFocus
    tmrUpdateTime.Enabled = True
    
End Sub

Private Sub pbTimeBlock_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    
    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    CalcPercent X
    frmPlayer.SetFocus
    tmrUpdateTime.Enabled = True
    
End Sub

Private Sub tmrUpdateTime_Timer()
    Sleep 2

    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    UpdateStatus mdlGlobalPlayer.Volume & "%, " & Round(mdlGlobalPlayer.Rate / 100, 2) & "x" & ", " & Format(Round(Precent, 2), "##.#0") & "% (" & FormatedCurrentTime & "/" & FormatedDuration & ")", PlayTime
    pbTimeBlock.width = Precent / 100 * (pbTimeBar.width)

    If (Duration > 1) Then
        If (Duration <= CurrentTime) Then
            tmrUpdateTime.Enabled = False
            
            If (mdlPlaylist.PlaylistPlayNext) Then
                CurrentTime = 0
                pbTimeBlock.width = 0
            Else
                
                If (frmMenu.mmStatus_Loop.Checked) Then
                    CurrentTime = 0
                    pbTimeBlock.width = 0
                    
                End If
                
            End If

        Else

            'resize
            ReCalcPlayWindow

            'AlphaHwnd bbMenuBar.hDC, frmPlayer.hDC, 150&, bbMenuBar.Width, bbMenuBar.Height
        End If
    
    End If

    'Set Me.Picture = mdlGlobalPlayer.NowFrame
End Sub
