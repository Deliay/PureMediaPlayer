VERSION 5.00
Begin VB.UserControl DirectView 
   BackColor       =   &H80000006&
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
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
   ScaleHeight     =   488
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   Begin VB.PictureBox diContenter 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7320
      Left            =   0
      ScaleHeight     =   488
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3495
      Begin PureMediaPlayer.DirectViewItem diList 
         Height          =   780
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1376
      End
   End
   Begin VB.VScrollBar vsPageControl 
      Height          =   7335
      Left            =   3495
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "DirectView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private objListItem      As DirectMap

Private colContorls      As Collection

Private lngLastIndex     As Long

Private lngLastMouseMove As Long

Public Event ListItemClick(ByVal Index As Long, ListItem As DirectViewItem)

Public Event ListItemDblClick(ByVal Index As Long, ListItem As DirectViewItem)

Public Function MouseWheelEvent(ByVal Direction As Boolean, ByVal Key As Integer)

End Function

Private Sub diList_onClick(Index As Integer)
    RaiseEvent ListItemClick(CLng(Index), diList(Index))
    Render

End Sub

Private Sub diList_onDblClick(Index As Integer)
    RaiseEvent ListItemDblClick(CLng(Index), diList(Index))
    Render

End Sub

Private Sub diList_onMouseMove(Index As Integer, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    If (lngLastMouseMove > -1) Then
        diList(lngLastMouseMove).onIdleStatus

        If (Not diList(Index).NowPlaying) Then diList(Index).onMoveStatus

    End If
    lngLastMouseMove = Index

End Sub

Private Sub UserControl_Initialize()
    Set objListItem = New DirectMap
    Set colContorls = New Collection
    lngLastIndex = -1
    lngLastMouseMove = -1

End Sub

Private Sub UserControl_Resize()
    Width = 3735
    vsPageControl.Height = Height
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colContorls.[_NewEnum]

End Property

Public Function AddItem(ByVal strPath As String, _
                        ByVal strLength As String) As DirectViewItem

    If (objListItem.Exist(strPath)) Then
        Set AddItem = ItemOf(strPath)
        Exit Function

    End If

    If (lngLastIndex = -1 Or diList.UBound = lngLastIndex) Then
        lngLastIndex = lngLastIndex + 1

        On Error Resume Next

        Load diList(lngLastIndex)

        On Error GoTo 0

    End If

    objListItem.AddKeyValue strPath, strLength
    diList(lngLastIndex).Path = strPath
    diList(lngLastIndex).Duration = strLength
    Set AddItem = diList(lngLastIndex)
    colContorls.Add diList(lngLastIndex)
    Render
    
    vsPageControl.Min = 0
    vsPageControl.Max = Fix((lngLastIndex + 1) / (Height / diList(0).Height))
    
    vsPageControl.Visible = (vsPageControl.Max <> 0)

End Function

Public Sub Render()

    Dim i             As Long

    Dim lngItemHeight As Long

    lngItemHeight = diList(0).Height

    For i = 0 To diList.UBound
        diList(i).Visible = False
    Next

    For i = 0 To lngLastIndex
        diList(i).Top = i * lngItemHeight

        If (diList(i).NowPlaying) Then
            diList(i).onSelectStatus
        Else
            diList(i).onIdleStatus
        End If

        diList(i).Visible = True
    Next

End Sub

Public Sub Clear()
    lngLastIndex = -1
    Render
    Set colContorls = New Collection
    Set objListItem = New DirectMap

End Sub

Public Property Get ItemOf(ByVal strPath As String) As DirectViewItem
Attribute ItemOf.VB_UserMemId = 0

    For Each ItemOf In colContorls

        If (ItemOf.Path = strPath) Then Exit Property
    Next

End Property

