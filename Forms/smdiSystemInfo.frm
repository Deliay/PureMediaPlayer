VERSION 5.00
Begin VB.Form frmSystemInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Associator"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
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
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "cmdClose"
      Height          =   360
      Left            =   4320
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   360
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   360
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.ListBox lstTypes 
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

Private Sub Form_Load()
<<<<<<< HEAD
    RaiseRegFileter Me.List1
=======
>>>>>>> master
    Me.ZOrder 0
    
End Sub

