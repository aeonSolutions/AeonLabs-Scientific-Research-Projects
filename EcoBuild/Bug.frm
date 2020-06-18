VERSION 5.00
Begin VB.Form Bug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report a Bug"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "Bug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label9 
      Caption         =   "or eMail to: "
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label8 
      Caption         =   "duracon@civil.uminho.pt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1140
      MouseIcon       =   "Bug.frx":324A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2400
      Width           =   1845
   End
   Begin VB.Label Label7 
      Caption         =   "PORTUGAL"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Department of Civil Engineering"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "University of Minho"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Campus de Azurém"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "4800-058 Guimarães"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "DURACON Software"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "If you have found a bug in this software, please report it in detail to:"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Bug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Label8_Click()
    ShellExecute Me.hwnd, vbNullString, "mailto:duracon@civil.uminho.pt", vbNullString, "C:\", 1
End Sub
