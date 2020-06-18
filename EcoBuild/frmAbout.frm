VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About [CODED]"
   ClientHeight    =   6495
   ClientLeft      =   4185
   ClientTop       =   3870
   ClientWidth     =   8280
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1620
      Picture         =   "frmAbout.frx":324A
      ScaleHeight     =   1665
      ScaleWidth      =   4545
      TabIndex        =   7
      Top             =   30
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "All Rights Reserved"
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
      Left            =   510
      TabIndex        =   11
      Top             =   2400
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
      Left            =   1950
      TabIndex        =   10
      Top             =   4440
      Width           =   3495
   End
   Begin VB.Label Title_four 
      Caption         =   "     Said Jalali, PhD"
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
      Left            =   690
      TabIndex        =   9
      Top             =   4110
      Width           =   2295
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
      Left            =   3360
      TabIndex        =   8
      Top             =   4740
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "    Alexandre Peyroteo, DEng"
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
      Left            =   690
      TabIndex        =   6
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "     Miguel Tomás Silva, DEng"
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
      Left            =   690
      TabIndex        =   5
      Top             =   3840
      Width           =   2955
   End
   Begin VB.Label Label2 
      Caption         =   "    Carla Carvalho, DEng"
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
      Left            =   690
      TabIndex        =   4
      Top             =   3540
      Width           =   3375
   End
   Begin VB.Label Disclaimer 
      Caption         =   $"frmAbout.frx":753A
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   270
      TabIndex        =   3
      Top             =   5280
      Width           =   7695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   150
      X2              =   7800
      Y1              =   5100
      Y2              =   5100
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
      Left            =   690
      TabIndex        =   2
      Top             =   2940
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   120
      X2              =   7680
      Y1              =   2790
      Y2              =   2790
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
      Left            =   120
      TabIndex        =   1
      Top             =   2100
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
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Me.Caption = App.Title
Title_one.Caption = App.Title & " " & App.Major & "." & App.Minor & " (" & build_date & ")"
Title_two.Caption = "Copyright © "
Disclaimer.Caption = App.Title & " and " & App.Title & " related softwares are intended for the use of professionals who are competent to evaluate the significance and limitations of its content and recommendations and who will accept responsibility for the application of the material it contains. The author responsible for the development of the program shall not be liable for any loss or damage arising there from. "
End Sub

