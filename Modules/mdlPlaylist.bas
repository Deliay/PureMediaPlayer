Attribute VB_Name = "mdlPlaylist"
Option Explicit
'
'Public Enum ConfigPart
'
'    Settings
'    History
'
'End Enum

Public colPlayItems  As Collection
 
Public strPlaylist   As String

Public playlistCount As Long

Public Function LoadPlaylist(ByVal strPath As String)
    frmMain.lstPlaylist.Clear
    strPlaylist = strPath
    Set colPlayItems = New Collection
    Open strPath For Input As #1
    
    Dim strItem As String, lngIter As Long
    
    Line Input #1, strItem
    
    playlistCount = val(strItem)
    
    For lngIter = 0 To playlistCount - 1
        
        Dim strFullPath As String, strLength As String

        If (EOF(1)) Then Exit For
        Line Input #1, strFullPath
        Line Input #1, strLength
        AddFileToPlaylist strFullPath, strLength
    Next
    
    Close #1
    
End Function

Public Function PlaylistPlayNext(Optional Prev As Boolean = False) As Boolean

    On Error GoTo ResumePlay

    PlaylistPlayNext = False

    Dim flag As Long

    If (GetItemIDByName(frmMain.NowPlaying.Title) = mdlPlaylist.colPlayItems.Count) Then Exit Function
    flag = 1

    If (Prev = True) Then flag = -1
    PlaylistPlayNext = PlayByName(colPlayItems(GetItemIDByName(frmMain.NowPlaying.Title) + flag).FullPath)

ResumePlay:

    mdlGlobalPlayer.CurrentTime = 0
    mdlGlobalPlayer.Play

End Function

Public Function GetItemIDByName(ByVal strName As String) As Long
    
    Dim lngIter As Long
    
    For lngIter = 1 To colPlayItems.Count
        
        If (strName = colPlayItems(lngIter).Name) Then
            GetItemIDByName = lngIter

            Exit Function
            
        End If
        
    Next
    
End Function

Public Function SetItemLength(strFullPath As String, Length As String)

    If (GetItemByPath(strFullPath) Is Nothing) Then Exit Function
    GetItemByPath(strFullPath).Length = Length
    frmMain.lstPlaylist(strFullPath).Duration = Length

    GlobalConfig.FileDuration.Value(MD5String(strFullPath)) = Length

    If (strPlaylist <> "") Then SavePlaylist
    
End Function

Public Function AddFileToPlaylist(ByVal strPath As String, _
                                  Optional Length As String = "") As Boolean
    
    Dim Item    As New PlayListItem
    
    Dim AddItem As DirectViewItem
    
    Item.FullPath = strPath
    Item.Name = NameGet(strPath)
    
    On Error GoTo notExist
    
    If (frmMain.lstPlaylist(Item.FullPath) Is Nothing) Then
notExist:
        
        On Error GoTo 0
        
        On Error GoTo Exist
        
        'Dim tmpFG As New FilgraphManager
        
        'Dim ifPOS As IMediaPosition
        
        'tmpFG.RenderFile strPath
        'Set ifPOS = tmpFG
        'item.Length = (ifPOS.Duration \ 60) & ":" & (ifPOS.Duration Mod 60)
        
        If (Not GlobalConfig.LastPlayList.Exist(Item.FullPath)) Then

            GlobalConfig.LastPlayList.AddItem Item.FullPath

        End If

        If (colPlayItems Is Nothing) Then Set colPlayItems = New Collection
        colPlayItems.Add Item, Item.FullPath
        playlistCount = playlistCount + 1
        Set AddItem = frmMain.lstPlaylist.AddItem(Item.FullPath, IIf(Length = "", GlobalConfig.FileDuration(MD5String(Item.FullPath)), Length))

        If (strPath = mdlGlobalPlayer.File) Then
            Set frmMain.NowPlaying = AddItem
            AddItem.Duration = mdlGlobalPlayer.FormatedDuration
            GetItemByPath(File).Length = AddItem.Duration

        End If
        
Exist:
        
        On Error GoTo 0
        
    Else
        
    End If
    
End Function

Public Function SavePlaylist()
    
    Dim varIter As Variant
    
    Open strPlaylist For Output As #1
    Print #1, playlistCount - 1
    
    For Each varIter In colPlayItems
        
        Print #1, varIter.FullPath
        Print #1, varIter.Length
    Next

    Close #1
    
End Function

Public Sub PlaylistClear()
    Set colPlayItems = New Collection
    frmMain.lstPlaylist.Clear
    
End Sub

Public Function PlayByName(ByVal strName As String) As Boolean
    frmMain.lstPlaylist_ListItemDblClick 0, frmMain.lstPlaylist(strName)
    
End Function

Public Function GetItemByPath(ByVal strPath As String) As PlayListItem

    Dim playItem As Variant

    For Each playItem In colPlayItems

        If (playItem.FullPath = strPath) Then
            Set GetItemByPath = playItem

            Exit Function

        End If

    Next
    
End Function
