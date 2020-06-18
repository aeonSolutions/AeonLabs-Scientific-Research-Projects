VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_perform_calculations 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Performing simulation"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   630
      TabIndex        =   0
      Top             =   930
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label txt 
      Caption         =   "task"
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
      Left            =   630
      TabIndex        =   1
      Top             =   660
      Width           =   5985
   End
End
Attribute VB_Name = "frm_perform_calculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

