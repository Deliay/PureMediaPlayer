VERSION 5.00
Begin VB.Form frmSystemInfo 
   AutoRedraw      =   -1  'True
   Caption         =   "App Register / File Associator"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "smdiSystemInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdUninstall 
      Caption         =   "Uninstall"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "Fix"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Remove All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox cbExtName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "smdiSystemInfo.frx":000C
      Left            =   120
      List            =   "smdiSystemInfo.frx":0022
      TabIndex        =   5
      Text            =   ".mp4"
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox lstTypes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label lblFileAssociator 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Associator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmSystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    If (GlobalConfig.BindedFileExts.Exist(cbExtName.Text)) Then
        MsgBox mdlLanguageApplyer.StaticString(EXT_ALREADY_BIND)

        Exit Sub
        
    Else

        Dim reg         As New RegisterEditor

        Dim strOldValue As String

        strOldValue = reg.GetString(HKEY_CLASSES_ROOT, cbExtName.Text, "")
        
        If Not (strOldValue = "PureMediaPlayer") Then

            'storage old value
            GlobalConfig.OldBindExts.AddKeyValue cbExtName.Text, strOldValue

        End If

        GlobalConfig.BindedFileExts.AddItem cbExtName.Text

        ReqAdminPerm "--bindext " & cbExtName.Text
    End If
    
    Me.lstTypes.AddItem cbExtName.Text
    mdlConfig.SaveConfig
End Sub

Private Sub cmdClose_Click()
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()

    If (lstTypes.Text = "") Then Exit Sub
    'revocer old setting
    ReqAdminPerm "--unbindext " & lstTypes.Text

    GlobalConfig.OldBindExts.Remove lstTypes.Text

    lstTypes.RemoveItem lstTypes.ListIndex

    GlobalConfig.BindedFileExts.TakeFrom lstTypes

    mdlConfig.SaveConfig
End Sub

Private Sub cmdFix_Click()
    ReqAdminPerm "--association"

    GlobalConfig.AppRegistered = "1"

    mdlConfig.SaveConfig
End Sub

Private Sub cmdRemoveAll_Click()
    lstTypes.Clear
    ReqAdminPerm "--unbindall"

    GlobalConfig.BindedFileExts.TakeTo lstTypes
    
End Sub

Private Sub cmdUninstall_Click()

    If (MsgBox(mdlLanguageApplyer.StaticString(TIPS_MAKSURE_UNINSTALL), vbYesNo, "Tips") = vbYes) Then
        lstTypes.Clear
        ReqAdminPerm "--uninstall"

        GlobalConfig.BindedFileExts.TakeTo lstTypes

        GlobalConfig.AppRegistered = "0"

    End If

    Unload Me
End Sub

Private Sub Form_Load()
        
    If (GlobalConfig.AppRegistered = "0") Then
        cmdFix_Click
    End If
    
    Me.ZOrder 0

    GlobalConfig.BindedFileExts.TakeTo Me.lstTypes

    'Load already associated ext form ini
    
    Me.cbExtName.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)

    GlobalConfig.BindedFileExts.TakeFrom Me.lstTypes

    mdlConfig.SaveConfig
    'Save association status to current config file
End Sub
