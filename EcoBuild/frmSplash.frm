VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3510
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7140
      TabIndex        =   1
      Top             =   2640
      Width           =   7140
      Begin VB.Label lblVersion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version Number"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   4380
         TabIndex        =   2
         Top             =   30
         Width           =   2595
      End
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   -30
      Picture         =   "frmSplash.frx":2052
      Top             =   0
      Width           =   7140
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Copyright © 2005 [CODED] - All Rights Reserved"
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   3150
      Width           =   6855
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    Label4.Caption = "Copyright © 2005 " & App.Title & " - All Rights Reserved"
End Sub

