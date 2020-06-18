VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3630
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7125
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
      TabIndex        =   3
      Top             =   2100
      Width           =   7140
      Begin VB.Label lblVersion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Version Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   4920
         TabIndex        =   4
         Top             =   60
         Width           =   2145
      End
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   0
      Picture         =   "frmSplash.frx":2052
      Top             =   0
      Width           =   7140
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Copyright © 2004 DURACON - All Rights Reserved"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   6855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Registered to Rui Miguel Ferreira (University of Minho)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Beta Edition - Freeware Mode (#0) - Serial # 1107.464.301.0181"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2760
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
End Sub

