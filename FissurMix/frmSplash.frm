VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4470
   ClientLeft      =   2970
   ClientTop       =   2715
   ClientWidth     =   7770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   -30
      Picture         =   "frmSplash.frx":08CA
      ScaleHeight     =   4545
      ScaleWidth      =   7785
      TabIndex        =   0
      Top             =   -60
      Width           =   7815
      Begin VB.Label lblversion 
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   4140
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2005 [CODED] - All Rights Reserved"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   2550
         TabIndex        =   1
         Top             =   4170
         Width           =   5085
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblversion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    Label1.Caption = "Copyright © 2005 " & App.Title & " - All Rights Reserved"
    lblversion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

