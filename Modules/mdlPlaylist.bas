Attribute VB_Name = "mdlPlaylist"
Option Explicit

Public Enum ConfigPart

    Settings
    History

End Enum

Public colPlayItems  As Collection
 
Public strPlaylist   As String

Public playlistCount As Long

Public Function LoadPlaylist(ByVal strPath As String)
    frmPlaylist.lstPlaylist.ListItems.Clear
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

Public Function PlaylistPlayNext() As Boolean
    On Error GoTo ResumePlay
    PlaylistPlayNext = False
    If (GetItemIDByName(frmPlaylist.nowPlaying.Text) = mdlPlaylist.colPlayItems.Count) Then Exit Function
    mdlGlobalPlayer.File = colPlayItems(GetItemIDByName(frmPlaylist.nowPlaying.Text) + 1).FullPath
    mdlGlobalPlayer.RenderMediaFile
    PlaylistPlayNext = True
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
    frmPlaylist.lstPlaylist.ListItems(strFullPath).SubItems(1) = Length
    
    If (strPlaylist <> "") Then SavePlaylist
    
End Function

Public Function AddFileToPlaylist(ByVal strPath As String, _
                                  Optional Length As String = "") As Boolean
    
    Dim Item    As New PlayListItem
    
    Dim addItem As ListItem
    
    Item.FullPath = strPath
    Item.Name = NameGet(strPath)
    
    On Error GoTo notExist
    
    If (frmPlaylist.lstPlaylist.ListItems(Item.FullPath) Is Nothing) Then
notExist:
        
        On Error GoTo 0
        
        On Error GoTo Exist
        
        'Dim tmpFG As New FilgraphManager
        
        'Dim ifPOS As IMediaPosition
        
        'tmpFG.RenderFile strPath
        'Set ifPOS = tmpFG
        'item.Length = (ifPOS.Duration \ 60) & ":" & (ifPOS.Duration Mod 60)
        If (colPlayItems Is Nothing) Then Set colPlayItems = New Collection
        colPlayItems.Add Item, Item.FullPath
        playlistCount = playlistCount + 1
        Set addItem = frmPlaylist.lstPlaylist.ListItems.Add(, Item.FullPath, Item.Name)
        addItem.SubItems(1) = Length
        
        If (strPath = mdlGlobalPlayer.File) Then
            addItem.Bold = True
            Set frmPlaylist.nowPlaying = addItem
            addItem.SubItems(1) = mdlGlobalPlayer.FormatedDuration
            GetItemByPath(File).Length = addItem.SubItems(1)
            
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
    frmPlaylist.lstPlaylist.ListItems.Clear
    
End Sub

Public Function PlayByName(ByVal strName As String) As Boolean
    mdlGlobalPlayer.File = GetItemByPath(strName).FullPath
    mdlGlobalPlayer.RenderMediaFile
    
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

Public Function getConfig(key As String, _
                          Optional ConfigType As ConfigPart = Settings) As String
    getConfig = GetSetting(App.ProductName, ConfigType, key)
    
End Function

Public Sub saveConfig(key As String, _
                      value As String, _
                      Optional ConfigType As ConfigPart = Settings)
    SaveSetting App.ProductName, ConfigType, key, value
    
End Sub
