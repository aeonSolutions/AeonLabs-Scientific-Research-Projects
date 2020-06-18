VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_perform_calculations 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Performing simulation"
   ClientHeight    =   1620
   ClientLeft      =   2670
   ClientTop       =   3450
   ClientWidth     =   7545
   Icon            =   "frm_perform_calculations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   840
      Width           =   5505
      _ExtentX        =   9710
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
      Left            =   1590
      TabIndex        =   1
      Top             =   510
      Width           =   5445
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   330
      Picture         =   "frm_perform_calculations.frx":08CA
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "frm_perform_calculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

