VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   Caption         =   "Pure Media Player"
   ClientHeight    =   6705
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox sbStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
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
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   0
         ScaleHeight     =   10
         ScaleMode       =   0  'User
         ScaleWidth      =   697
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
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
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
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
         Left            =   4200
         TabIndex        =   13
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0:00/0:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   3000
         TabIndex        =   12
         Top             =   360
         Width           =   1050
      End
   End
   Begin MSComctlLib.ListView lstPlaylist 
      Height          =   6255
      Left            =   10440
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   11033
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4941
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   1235
      EndProperty
   End
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   0
      ScaleHeight     =   417
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   697
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10455
      Begin VB.PictureBox bbPlaylist 
         Appearance      =   0  'Flat
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

Public nowPlaying    As ListItem

Public isHide        As Boolean

Private ItemSelected As ListItem

Public srcH As Long, srcW As Long

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
        
        If (Not mdlToolBarAlphaer.boolPlaylistStatus) Then
            SwitchUI True

        End If

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

Private Sub Form_Activate()
    mdlGlobalPlayer.SwitchFullScreen True, False
    RenderUI
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
    UpdateStatus StaticString(PLAYER_STATUS_READY), Action
    UpdateStatus StaticString(PLAY_STATUS_STOPED), PlayBack
    UpdateStatus StaticString(FILE_STATUS_NOFILE), StatusBarEnum.FileName
    SwitchUI True
    RenderUI
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

Public Sub Form_Resize()
    ReCalcPlayWindow
    If (mdlToolBarAlphaer.UIStatus = True) Then RefreshUI
End Sub

Public Sub ReCalcPlayWindow()
    frmPlayer.Width = (Me.Width / Screen.TwipsPerPixelX)
    frmPlayer.Height = (Me.Height / Screen.TwipsPerPixelY)

    bbPlaylist.Left = frmPlayer.Width - 32 - mdlToolBarAlphaer.UIWidthRight
    bbPlaylist.Top = frmPlayer.Height / 2 - bbPlaylist.Height
    lstPlaylist.Left = frmPlayer.Width - mdlToolBarAlphaer.UIWidthRight
    
    frmPlayer.Width = (Me.Width / Screen.TwipsPerPixelX) - mdlToolBarAlphaer.UIWidthRight
    mdlGlobalPlayer.Width = frmPlayer.Width
    mdlGlobalPlayer.Height = frmPlayer.Height - mdlToolBarAlphaer.UIHeightButtom - 23
    
    lstPlaylist.Height = mdlGlobalPlayer.Height
    
    pbTimeBar.Width = (Me.Width / Screen.TwipsPerPixelX)
    
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
        mdlToolBarAlphaer.apMenuButton.RefreshHW 32, 32

    End If

    If (bbPlaylist.BackColor <> 0) Then
        bbPlaylist.BackColor = 0
        mdlToolBarAlphaer.apPlaylistHint.RefreshHW 24, 48

    End If

    If (X < 32 And Y < 32) Then
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
            'AlphaHwnd bbMenuBar.hDC, frmPlayer.hDC, 150&, bbMenuBar.Width, bbMenuBar.Height
        End If
    
    End If

    'Set Me.Picture = mdlGlobalPlayer.NowFrame
End Sub

Public Sub AutoPatern()
    Load frmPaternAdd
    frmPaternAdd.Show vbModeless, Me
    frmPaternAdd.Form_Load
    frmPaternAdd.cmdAddToList_Click
    
End Sub

Private Sub lstPlaylist_DblClick()
    
    If (ItemSelected Is Nothing) Then Exit Sub
    If (Not (NameGet(File) = ItemSelected.Text)) Then
        mdlGlobalPlayer.CloseFile
        mdlGlobalPlayer.File = mdlPlaylist.GetItemByPath(ItemSelected.key).FullPath
        mdlGlobalPlayer.RenderMediaFile
        
        If Not nowPlaying Is Nothing Then nowPlaying.Bold = False
        Set nowPlaying = ItemSelected
        
        If Not nowPlaying Is Nothing Then nowPlaying.Bold = True
        
    End If
    
End Sub

Private Sub lstPlaylist_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set ItemSelected = Item
    
End Sub

Private Sub lstPlaylist_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = vbKeySpace) Then SwitchPlayStauts
    
End Sub

