VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Pure Media Player"
   ClientHeight    =   6705
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pmdiMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox sbStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5850
      Width           =   10455
      Begin VB.PictureBox bbPlaystatus 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   2520
         MousePointer    =   7  'Size N S
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox bbPlaystatus 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox bbPlaystatus 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   1920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox bbPlaystatus 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox bbPlaystatus 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   720
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.PictureBox pbTimeBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   0
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   697
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   10455
         Begin VB.PictureBox pbTimeBlock 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   0
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   1
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Stop"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3120
         TabIndex        =   13
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0:00/0:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3960
         TabIndex        =   12
         Top             =   360
         Width           =   2970
      End
   End
   Begin vbalListViewLib6.vbalListViewCtl lstPlaylist 
      Height          =   6255
      Left            =   10440
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11033
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      BackColor       =   1842204
      View            =   1
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      FullRowSelect   =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      FlatScrollBar   =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      DoubleBuffer    =   -1  'True
      NoColumnHeaders =   -1  'True
      TileBackgroundPicture=   0   'False
   End
   Begin VB.PictureBox bbMenuBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   480
   End
   Begin VB.Timer tmrUpdateTime 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   2760
   End
   Begin VB.PictureBox frmPlayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10455
      Begin VB.Timer tmrTextRender 
         Enabled         =   0   'False
         Left            =   3720
         Top             =   2760
      End
      Begin VB.PictureBox bbPlaylist 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   9840
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   24
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2880
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

Public nowPlaying      As cListItem

Public isHide          As Boolean

Public mouseDownStatus As Boolean

Public dragY           As Single

Private ItemSelected   As cListItem

Public srcH            As Long, srcW As Long

Private Sub asd_Click()

End Sub

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
        mdlToolBarAlphaer.apMenuButton.hDC = bbMenuBar.hDC
        mdlToolBarAlphaer.apMenuButton.RefreshHW 32, 32
        SwitchUI True

        If (boolPlaylistStatus = True) Then
            PlaylistShow
        Else
            PlaylistHide

            'RefreshUI
        End If

    End If

End Sub

Private Sub bbPlaylist_Click()

    If (boolPlaylistStatus = True) Then
        PlaylistHide
    Else
        PlaylistShow

        'RefreshUI
    End If
    
End Sub

Private Sub bbPlaylist_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    
    If (bbPlaylist.BackColor <> RGB(48, 48, 48)) Then
        bbPlaylist.BackColor = RGB(48, 48, 48)
        mdlToolBarAlphaer.apPlaylistHint.hDC = bbPlaylist.hDC
        mdlToolBarAlphaer.apPlaylistHint.RefreshHW 24, 48

    End If

    If (Not mdlToolBarAlphaer.boolPlaylistStatus) Then
        SwitchUI True

    End If

End Sub

Private Sub bbPlaystatus_Click(Index As Integer)

    Dim clickType As PlayControl

    clickType = Index
    
    Select Case clickType

        Case PlayControl.CTRL_PLAYPAUSE
            'PlayPauseSwitch
            mdlGlobalPlayer.SwitchPlayStauts

        Case PlayControl.CTRL_STOP
            mdlGlobalPlayer.StopPlay

        Case PlayControl.CTRL_NEXT
            mdlPlaylist.PlaylistPlayNext

        Case PlayControl.CTRL_PREV
            mdlPlaylist.PlaylistPlayNext True
        
        Case PlayControl.CTRL_VOICE
            
    End Select

    frmPlayer.SetFocus

End Sub

Private Sub bbPlaystatus_MouseDown(Index As Integer, _
                                   Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
    mouseDownStatus = True
    dragY = Y

End Sub

Private Sub bbPlaystatus_MouseMove(Index As Integer, _
                                   Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)

    Dim lngFlag As Long

    If (mouseDownStatus And Index = PlayControl.CTRL_VOICE And mdlGlobalPlayer.GlobalPlayStatus = Running And mdlGlobalPlayer.Loaded) Then
        If (dragY > Y) Then

            'drag up
            'voice up
            If (dragY - Y > 2) Then
                lngFlag = 1

            End If

        ElseIf (dragY < Y) Then

            If (Y - dragY > 2) Then
                lngFlag = -1

            End If

        End If

        dragY = Y

        If (lngFlag <> 0) Then
            mdlGlobalPlayer.Volume = mdlGlobalPlayer.Volume + lngFlag
            
            lngFlag = 0

        End If

    End If

End Sub

Private Sub bbPlaystatus_MouseUp(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    mouseDownStatus = False
    dragY = 0
    
End Sub

Private Sub Form_Activate()
    mdlGlobalPlayer.SwitchFullScreen True, False

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If (22 = KeyAscii) Then
        If (OpenClipboard(Me.hWnd)) Then
            mdlMsgProc.ReadDrapQueryFile GetClipboardData(CF_HDROP)
            CloseClipboard
            
        End If
    
    ElseIf (4 = KeyAscii) Then
        frmMenu.mmFile_Open_Click
        
    ElseIf (3 = KeyAscii) Then
        frmMenu.mmFile_Close_Click
        
    ElseIf (24 = KeyAscii) Then
        frmMenu.mmFile_Exit_Click
        
    ElseIf (6 = KeyAscii) Then
        frmMenu.mmStatus_ShowPlaylist_Click
    ElseIf (16 = KeyAscii) Then
        frmPaternAdd.Show vbModal, Me
    ElseIf (20 = KeyAscii) Then
        frmMenu.mmFile_LoadSubtitle_Click
    ElseIf (26 = KeyAscii) Then
        mdlToolBarAlphaer.ShowText InputBox("test")
    End If

End Sub

Private Sub frmPlayer_Click()
    SwitchUI

    If (mdlToolBarAlphaer.UIStatus = False) Then PlaylistHide

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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

Private Sub Form_Load()
    Me.Show
    lstPlaylist.Columns.Add , , "File", , 2400
    lstPlaylist.Columns.Add , , "Duration", , 800
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    UpdateStatus StaticString(FILE_STATUS_NOFILE), StatusBarEnum.FileName
    GlobalConfig.LastHwnd = CStr(Me.hWnd)
    mdlConfig.SaveConfig
    Me.Show
    frmPlayer.AutoRedraw = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeWindow Me.hWnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bbMenuBar.BackColor = vbBlack

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    ExitProgram
    
End Sub

Public Sub Form_Resize()
    ReCalcPlayWindow

    If (mdlToolBarAlphaer.UIStatus = True) Then RefreshUI

End Sub

Public Sub ReCalcPlayWindow()

    If (Me.Width = 0) Then Exit Sub
    If (Me.Height = 0) Then Exit Sub
    
    frmPlayer.Width = Me.ScaleWidth
    frmPlayer.Height = Me.ScaleHeight

    If (frmPlayer.Width <= 160) Then Exit Sub
    If (frmPlayer.Height <= 100) Then Exit Sub

    bbPlaylist.Left = frmPlayer.Width - 24 - mdlToolBarAlphaer.UIWidthRight
    bbPlaylist.Top = frmPlayer.Height / 2 - bbPlaylist.Height
    lstPlaylist.Left = frmPlayer.Width - mdlToolBarAlphaer.UIWidthRight
    
    frmPlayer.Width = Me.ScaleWidth - mdlToolBarAlphaer.UIWidthRight
    'frmPlayer.Width = pbTimeBar.Top
    mdlGlobalPlayer.Width = frmPlayer.Width
    mdlGlobalPlayer.Height = frmPlayer.Height - mdlToolBarAlphaer.UIHeightButtom
    
    lstPlaylist.Height = mdlGlobalPlayer.Height
    
    pbTimeBar.Width = Me.ScaleWidth

    ResizePlayWindow

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    ExitProgram
    
End Sub

Private Sub frmPlayer_DblClick()
    SwitchFullScreen
    SwitchUI True

End Sub

Private Sub frmPlayer_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    
    If (Button = vbRightButton) Then Me.PopupMenu frmMenu.mmStatus

End Sub

Private Sub CalcPercent(X As Single)
    Precent = (X / pbTimeBar.Width) * 100
    tmrUpdateTime_Timer
    
End Sub

Private Sub frmPlayer_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    If (bbMenuBar.BackColor <> 0) Then
        bbMenuBar.BackColor = 0
        mdlToolBarAlphaer.apMenuButton.hDC = bbMenuBar.hDC
        mdlToolBarAlphaer.apMenuButton.RefreshHW 32, 32

    End If

    If (bbPlaylist.BackColor <> 0) Then
        bbPlaylist.BackColor = 0
        mdlToolBarAlphaer.apPlaylistHint.hDC = bbPlaylist.hDC
        mdlToolBarAlphaer.apPlaylistHint.RefreshHW 24, 48

    End If

    If (X < 32 And Y < 32) Then
        RefreshUI

    End If

End Sub

Public Sub lstPlaylist_ItemDblClick(Item As vbalListViewLib6.cListItem)

    If (Not (NameGet(mdlGlobalPlayer.File) = Item.Text)) Then
        mdlGlobalPlayer.CloseFile
        mdlGlobalPlayer.File = mdlPlaylist.GetItemByPath(Item.Key).FullPath
        mdlGlobalPlayer.RenderMediaFile
        
        If Not nowPlaying Is Nothing Then nowPlaying.ForeColor = vbWhite
        Set nowPlaying = Item
        nowPlaying.ForeColor = vbGrayText
        
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

Private Sub tmrTextRender_Timer()
    tmrTextRender.Enabled = False
    frmPlayer.Line (32, 0)-(frmPlayer.ScaleWidth, 16), vbBlack, BF
    
End Sub

Private Sub tmrUpdateTime_Timer()
    Sleep 1

    If (mdlGlobalPlayer.Loaded = False) Then Exit Sub
    UpdateStatus FormatedCurrentTime & "/" & FormatedDuration, PlayTime
    pbTimeBlock.Width = Precent / 100 * (pbTimeBar.Width)

    If (Duration > 1) Then
        If (Duration <= CurrentTime) Then
            tmrUpdateTime.Enabled = False
            
            If (mdlPlaylist.PlaylistPlayNext) Then
                CurrentTime = 0
                pbTimeBlock.Width = 0
            Else
                
                If (frmMenu.mmStatus_Loop.Checked) Then
                    CurrentTime = 0
                    pbTimeBlock.Width = 0
                    
                End If
                
            End If

        Else

            'resize
            If (srcH <> Me.Height Or srcW <> Me.Width) Then
                Form_Resize
                srcH = Me.Height
                srcW = Me.Width

            End If

        End If
    
    End If


    If (mdlGlobalPlayer.GlobalRenderType = EnhancedVideoRenderer) Then
        If (EVRHoster.srcT > 0) Then
            frmPlayer.Line (0, 0)-(frmMain.frmPlayer.ScaleWidth, EVRHoster.srcT - 1), vbBlack, BF
            frmPlayer.Line (0, EVRHoster.srcH + EVRHoster.srcT)-(frmPlayer.ScaleWidth, frmPlayer.ScaleHeight), vbBlack, BF
        End If
        If (EVRHoster.srcL > 0) Then
            frmPlayer.Line (0, 0)-(EVRHoster.srcL - 1, frmPlayer.ScaleHeight), vbBlack, BF
            frmPlayer.Line (EVRHoster.srcL + EVRHoster.srcW, 0)-(frmPlayer.ScaleWidth, frmPlayer.ScaleHeight), vbBlack, BF
        End If
    End If
End Sub

Public Sub AutoPatern()
    Load frmPaternAdd
    frmPaternAdd.Show vbModeless, Me
    frmPaternAdd.Form_Load
    frmPaternAdd.cmdAddToList_Click
    
End Sub


