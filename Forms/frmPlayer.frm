VERSION 5.00
Begin VB.Form frmPlayer 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Idle"
   ClientHeight    =   4770
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   318
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   438
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    
    'SwitchPlayStauts
End Sub

Private Sub Form_DblClick()
    SwitchFullScreen
    
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (Button = vbRightButton) Then Me.PopupMenu frmMain.mmStatus
    
End Sub
