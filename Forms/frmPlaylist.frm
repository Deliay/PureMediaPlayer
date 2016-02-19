VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPlaylist 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Playlist"
   ClientHeight    =   5025
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   3735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlaylist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMovePadder 
      Interval        =   500
      Left            =   1320
      Top             =   2280
   End
   Begin MSComctlLib.ListView lstPlaylist 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8705
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
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
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu File_Add 
         Caption         =   "Add File"
         Shortcut        =   ^A
      End
      Begin VB.Menu File_Spec2 
         Caption         =   "-"
      End
      Begin VB.Menu File_PaternFind 
         Caption         =   "Patern Find"
         Enabled         =   0   'False
      End
      Begin VB.Menu File_Spec1 
         Caption         =   "-"
      End
      Begin VB.Menu File_New 
         Caption         =   "New"
      End
      Begin VB.Menu File_Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu File_Load 
         Caption         =   "Load"
      End
      Begin VB.Menu File_Clear 
         Caption         =   "Clear"
      End
      Begin VB.Menu File_Spce2 
         Caption         =   "-"
      End
      Begin VB.Menu File_Close 
         Caption         =   "Close"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu Playlist 
      Caption         =   "Playlist"
      Begin VB.Menu Playlist_Collection 
         Caption         =   "(None)"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu ItemAction 
      Caption         =   "ItemAction"
      Visible         =   0   'False
      Begin VB.Menu ItemAction_Remove 
         Caption         =   "Remove"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu ItemAction_Play 
         Caption         =   "Play"
      End
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nowPlaying    As ListItem

Public isHide        As Boolean

Private ItemSelected As ListItem

Private oldTop       As Long

Private oldLeft      As Long

Private Sub File_Add_Click()
    cdlg.FileName = ""
    cdlg.ShowOpen
    AddFileToPlaylist cdlg.FileName
    
End Sub

Private Sub File_Clear_Click()
    Set mdlPlaylist.colPlayItems = New Collection
    lstPlaylist.ListItems.Clear
    
End Sub

Private Sub File_Close_Click()
    frmPlaylist.isHide = Not frmPlaylist.isHide
    
    If (frmPlaylist.isHide) Then
        frmPlaylist.Show vbModeless, frmMain
    Else
        frmPlaylist.Hide
        
    End If
    
End Sub

Private Sub File_Load_Click()
    cdlg.FileName = ""
    cdlg.ShowOpen "PmP播放列表 (*.PPL)" & vbNullChar & "*.PPL"
    strPlaylist = cdlg.FileName
    
    If ((Len(strPlaylist) > 7) And (Right(strPlaylist, 3) = "PPL")) Then LoadPlaylist strPlaylist
    
End Sub

Private Sub File_New_Click()
    File_Clear_Click
    mdlPlaylist.strPlaylist = ""
    
End Sub

Private Sub File_PaternFind_Click()
    frmPaternAdd.Show
    
End Sub

Public Sub AutoPatern()
    frmPaternAdd.Show vbModeless, Me
    frmPaternAdd.cmdAddToList_Click
    
End Sub

Private Sub File_Save_Click()
    
    If (strPlaylist = "" Or Dir(strPlaylist) = "") Then
        cdlg.FileName = ""
        cdlg.ShowSave "PmP播放列表 (*.PPL)" & vbNullChar & "*.PPL"
        
        If ((Len(cdlg.FileName) > 7) And (Right(cdlg.FileName, 3) = "PPL")) Then
            mdlPlaylist.strPlaylist = cdlg.FileName
            SavePlaylist
            
        End If
        
    End If
    
End Sub

Private Sub Form_Load()
    tmrMovePadder_Timer
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    File_Close_Click
    
End Sub

Private Sub lstPlaylist_DblClick()
    
    If (ItemSelected Is Nothing) Then Exit Sub
    If (Not (NameGet(File) = ItemSelected.Text)) Then
        mdlGlobalPlayer.CloseFile
        mdlGlobalPlayer.LoadMediaFile mdlPlaylist.colPlayItems(ItemSelected.key).FullPath
        
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

Private Sub tmrMovePadder_Timer()
    
    If (frmMain.Left = oldLeft) Then Exit Sub
    If (frmMain.Top = oldTop) Then Exit Sub
    frmPlaylist.Top = frmMain.Top + frmPlaylist.Height / 2 - (35 * Screen.TwipsPerPixelY)
    frmPlaylist.Left = frmMain.Left - frmPlaylist.Width
    
    If (frmPlaylist.Left < 0) Then
        frmPlaylist.Left = frmMain.Left + frmMain.Width
        
    End If
    
    If (frmPlaylist.Left > Screen.Width) Then
        frmPlaylist.Left = frmMain.Left + frmMain.Width - frmPlaylist.Width
        
    End If
    
    oldLeft = frmMain.Left
    oldTop = frmMain.Top
    DoEvents
    
End Sub
