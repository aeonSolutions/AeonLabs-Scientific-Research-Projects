VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About DURACON"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   1770
      Picture         =   "frmAbout.frx":324A
      ScaleHeight     =   1395
      ScaleWidth      =   4755
      TabIndex        =   6
      Top             =   90
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "     Miguel Tomás Silva"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3870
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "Graphical User Interface (GUI)  by"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   3615
   End
   Begin VB.Label Disclaimer 
      Caption         =   $"frmAbout.frx":6426
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   270
      TabIndex        =   7
      Top             =   4500
      Width           =   7695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   30
      X2              =   7680
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Label Title_six 
      Caption         =   "    University of Minho - Portugal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   5
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Title_four 
      Caption         =   "     Miguel Ferreira, PhD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3270
      Width           =   2295
   End
   Begin VB.Label Title_three 
      Caption         =   "Developed by"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2970
      Width           =   2895
   End
   Begin VB.Label Title_five 
      Caption         =   "    Department of Civil Engineering"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2220
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   90
      X2              =   7650
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label Title_two 
      Caption         =   "Copyright (c) 2005  - All Rights Reserved"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   4365
   End
   Begin VB.Label Title_one 
      Caption         =   "DURACON (20040322)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Title_one.Caption = "DURACON " & App.Major & "." & App.Minor & " (20041117)"
End Sub

