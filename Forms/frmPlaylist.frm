VERSION 5.00
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
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrMovePadder 
      Interval        =   250
      Left            =   1200
      Top             =   2280
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
    frmMain.isHide = Not frmMain.isHide
    
    If (frmMain.isHide) Then
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

Private Sub tmrMovePadder_Timer()

    If (frmMain.Left = oldLeft) Then Exit Sub
    If (frmMain.Top = oldTop) Then Exit Sub
    frmPlaylist.Top = frmMain.Top + ((frmMain.height - frmPlaylist.height))
    frmPlaylist.Left = frmMain.Left - frmPlaylist.width
    
    If (frmPlaylist.Left < 0) Then
        frmPlaylist.Left = frmMain.Left + frmMain.width
        
    End If
    
    If (frmPlaylist.Left > Screen.width) Then
        frmPlaylist.Left = frmMain.Left + frmMain.width - frmPlaylist.width
        
    End If
    
    oldLeft = frmMain.Left
    oldTop = frmMain.Top
    DoEvents
    
End Sub
