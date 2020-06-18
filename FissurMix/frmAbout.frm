VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About [CODED]"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   990
      Left            =   90
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   150
      Width           =   1560
   End
   Begin VB.Label Label4 
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
      Left            =   180
      TabIndex        =   6
      Top             =   1920
      Width           =   3495
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
      Left            =   2790
      MouseIcon       =   "frmAbout.frx":1E2A
      TabIndex        =   8
      Top             =   2490
      Width           =   3015
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
      Height          =   225
      Left            =   180
      TabIndex        =   5
      Top             =   1680
      Width           =   2955
   End
   Begin VB.Label Disclaimer 
      Caption         =   $"frmAbout.frx":2134
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
      TabIndex        =   3
      Top             =   3060
      Width           =   7695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   150
      X2              =   7800
      Y1              =   2850
      Y2              =   2850
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
      TabIndex        =   2
      Top             =   1380
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   7680
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Title_two 
      Caption         =   "Copyright © [CODED]"
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
      Left            =   1650
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Title_one 
      Caption         =   "[CODED]  ([CODED])"
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
      Left            =   1650
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label2 
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
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Top             =   2190
      Width           =   3375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Form_Load()
Me.Caption = App.Title
Title_one.Caption = App.Title & " " & App.Major & "." & App.Minor & " (" & build_date & ")"
Title_two.Caption = "Copyright © "
Disclaimer.Caption = App.Title & " and " & App.Title & " related softwares are intended for the use of professionals who are competent to evaluate the significance and limitations of its content and recommendations and who will accept responsibility for the application of the material it contains. The author responsible for the development of the program shall not be liable for any loss or damage arising there from. "
End Sub

Private Sub Label1_Click()
    ShellExecute Me.hwnd, vbNullString, "http://www.miguelsilva.web.pt", vbNullString, "C:\", 1
End Sub
