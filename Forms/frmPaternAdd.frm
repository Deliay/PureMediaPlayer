VERSION 5.00
Begin VB.Form frmPaternAdd 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PaternAdd"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaternAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   236
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   4680
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdRemoveItem 
      Caption         =   "Remove Select"
      Height          =   360
      Left            =   3000
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddToList 
      Caption         =   "Add All To Playlist"
      Height          =   360
      Left            =   1320
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CheckBox chkEnableRegExp 
      Caption         =   "RegExp"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdTestParten 
      Caption         =   "TestView"
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   990
   End
   Begin VB.Frame ffView 
      Caption         =   "View"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   5655
      Begin VB.ListBox lstTestView 
         Height          =   1425
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.TextBox txtRegExp 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtNowName 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   4695
   End
   Begin VB.Label lblRegExTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patern"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   405
   End
End
Attribute VB_Name = "frmPaternAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private filesCol As Collection

Private rx       As New RegExp

Public Sub cmdAddToList_Click()
    
    Dim varIter As Variant
    
    For Each varIter In filesCol
        
        AddFileToPlaylist varIter
        
    Next
    Unload Me
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdRemoveItem_Click()
    
    Dim i As Long
    
    For i = lstTestView.ListCount - 1 To 0 Step -1
        
        If (lstTestView.Selected(i)) Then lstTestView.RemoveItem i
    Next
    
End Sub

Private Sub cmdTestParten_Click()
    
    Dim strFile As String
    
    Set filesCol = New Collection
    lstTestView.Clear
    
    If (chkEnableRegExp.Value = 0) Then
        'use default pattern
        strFile = Dir(DirGet(File) & "\" & txtRegExp.Text)
        
        While strFile <> ""
            
            lstTestView.AddItem strFile
            filesCol.Add DirGet(File) & "\" & strFile
            strFile = Dir
        Wend
    Else
        rx.MultiLine = False
        rx.IgnoreCase = True
        rx.Pattern = txtRegExp.Text
        
        On Error GoTo ErrExp
        
        rx.Test strFile
        
        On Error GoTo 0
        
        strFile = Dir(DirGet(File) & "\*.*")
        
        While strFile <> ""
            
            If (rx.Test(strFile)) Then
                lstTestView.AddItem strFile
                filesCol.Add DirGet(File) & "\" & strFile
                
            End If
            
            strFile = Dir()
        Wend
        Exit Sub
ErrExp:
        txtRegExp.Text = "Error Regular Expression!"
        
    End If
    
End Sub

Public Sub Form_Load()

    If (mdlGlobalPlayer.File = "") Then Exit Sub
    txtNowName.Text = NameGet(mdlGlobalPlayer.File)
    txtRegExp.Text = "*" & Mid$(NameGet(mdlGlobalPlayer.File), InStrRev(NameGet(mdlGlobalPlayer.File), "."))
    cmdTestParten_Click

End Sub

