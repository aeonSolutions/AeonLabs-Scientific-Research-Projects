VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Temperus"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "About Temperus"
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   990
      Left            =   60
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1560
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4290
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   2985
      Width           =   1467
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "http://www.miguelsilva.web.pt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1410
      MouseIcon       =   "frmAbout.frx":1E2A
      TabIndex        =   6
      Top             =   2550
      Width           =   3015
   End
   Begin VB.Label lblDescription 
      Caption         =   "Statistic analisys of Craking in stabilized ZrO2 coattings "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   60
      TabIndex        =   5
      Tag             =   "App Description"
      Top             =   1380
      Width           =   5715
   End
   Begin VB.Label lblTitle 
      Caption         =   "FissurMax"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1680
      TabIndex        =   4
      Tag             =   "Application Title"
      Top             =   240
      Width           =   3945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   5672
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   270
      X2              =   5687
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1680
      TabIndex        =   3
      Tag             =   "Version"
      Top             =   1050
      Width           =   3735
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "CopyRight 2004"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   60
      TabIndex        =   2
      Tag             =   "Warning: ..."
      Top             =   3090
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub form_load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub



Private Sub cmdOK_Click()
        Unload Me
End Sub

Private Sub Label1_Click()
    ShellExecute Me.hwnd, vbNullString, "http://www.miguelsilva.web.pt", vbNullString, "C:\", 1

End Sub
