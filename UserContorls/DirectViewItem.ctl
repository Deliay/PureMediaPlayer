VERSION 5.00
Begin VB.UserControl DirectViewItem 
   BackColor       =   &H80000006&
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   Begin VB.Label lblPlayStatus 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   465
      Width           =   1500
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1740
      TabIndex        =   1
      Top             =   480
      Width           =   1620
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3225
   End
End
Attribute VB_Name = "DirectViewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event onMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event onMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event onMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Event onDblClick()

Public Event onClick()

Private boolIsPlay As Boolean

Private strPath    As String

Private Sub lblCaption_Click()
    RaiseEvent onClick

End Sub

Private Sub lblCaption_DblClick()
    RaiseEvent onDblClick

End Sub

Private Sub lblTime_Click()
    RaiseEvent onClick

End Sub

Private Sub lblTime_DblClick()
    RaiseEvent onDblClick

End Sub

Private Sub UserControl_Click()
    RaiseEvent onClick

End Sub

Private Sub UserControl_DblClick()
    RaiseEvent onDblClick

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    RaiseEvent onMouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    RaiseEvent onMouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    RaiseEvent onMouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Resize()
    Width = 3495
    Height = 780

End Sub

Public Property Let Path(ByVal Value As String)
    strPath = Value
    lblCaption = NameGet(Value)

End Property

Public Property Get Path() As String
    Path = strPath

End Property

Public Property Let Title(ByVal Value As String)
    lblCaption.Caption = Value

End Property

Public Property Get Title() As String
    Title = lblCaption.Caption

End Property

Public Property Let Duration(ByVal Value As String)
    lblTime.Caption = Value

End Property

Public Property Get Duration() As String
    Duration = lblTime.Caption

End Property

Public Property Get NowPlaying() As Boolean
    NowPlaying = boolIsPlay

End Property

Public Property Let NowPlaying(Value As Boolean)
    boolIsPlay = Value

    If (boolIsPlay = True) Then
        onSelectStatus
        lblPlayStatus.Caption = "ÕýÔÚ²¥·Å"
    Else
        onIdleStatus
        lblPlayStatus.Caption = ""

    End If

End Property

Public Sub onMoveStatus()
    BackColor = &H80000001
    ForeColor = vbBlack

End Sub

Public Sub onSelectStatus()
    BackColor = vbGrayed
    ForeColor = vbBlack

End Sub

Public Sub onIdleStatus()
    BackColor = &H80000006
    ForeColor = vbWhite

End Sub
