VERSION 5.00
Begin VB.Form frm_concrete_board2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Beams"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9870
   Icon            =   "frm_concrete_board2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   375
      Left            =   8610
      TabIndex        =   22
      Top             =   5340
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "[CODED]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   3960
      TabIndex        =   10
      Top             =   210
      Width           =   5775
      Begin VB.TextBox costs_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4020
         LinkItem        =   "costs_txt"
         TabIndex        =   52
         Top             =   900
         Width           =   1005
      End
      Begin VB.ComboBox struct_type 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Insert"
         Height          =   315
         Left            =   210
         TabIndex        =   42
         Top             =   4140
         Width           =   5295
      End
      Begin VB.TextBox f5_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   8
         Text            =   "0"
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox f6_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2010
         TabIndex        =   9
         Text            =   "0"
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox f8_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   51
         Text            =   "0"
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox f10_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4860
         TabIndex        =   11
         Text            =   "0"
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox f12_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   12
         Text            =   "0"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox f16_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2010
         TabIndex        =   13
         Text            =   "0"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox f20_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Text            =   "0"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox f25_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4860
         TabIndex        =   15
         Text            =   "0"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox f32_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   540
         TabIndex        =   16
         Text            =   "0"
         Top             =   3660
         Width           =   735
      End
      Begin VB.TextBox num_beams 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4020
         TabIndex        =   7
         Top             =   1860
         Width           =   1005
      End
      Begin VB.TextBox lenght_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4020
         TabIndex        =   6
         Top             =   1380
         Width           =   1005
      End
      Begin VB.TextBox width_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   5
         Top             =   1620
         Width           =   705
      End
      Begin VB.TextBox height_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1110
         TabIndex        =   4
         Top             =   1170
         Width           =   705
      End
      Begin VB.Label Label33 
         Caption         =   "Costs"
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
         Left            =   3420
         TabIndex        =   54
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label32 
         Caption         =   "€/m3"
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
         Left            =   5130
         TabIndex        =   53
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label31 
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   50
         Top             =   1680
         Width           =   435
      End
      Begin VB.Label Label30 
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1860
         TabIndex        =   49
         Top             =   1200
         Width           =   435
      End
      Begin VB.Label Label29 
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5070
         TabIndex        =   48
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label25 
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   43
         Top             =   390
         Width           =   585
      End
      Begin VB.Label Label10 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   32
         Top             =   2610
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1830
         TabIndex        =   31
         Top             =   2610
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3300
         TabIndex        =   30
         Top             =   2610
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         TabIndex        =   29
         Top             =   2610
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   300
         TabIndex        =   28
         Top             =   3150
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1740
         TabIndex        =   27
         Top             =   3150
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3210
         TabIndex        =   26
         Top             =   3150
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   25
         Top             =   3180
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   270
         TabIndex        =   24
         Top             =   3720
         Width           =   375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000E&
         X1              =   240
         X2              =   5610
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label24 
         Caption         =   "Steel bars"
         Height          =   315
         Left            =   210
         TabIndex        =   23
         Top             =   2190
         Width           =   1125
      End
      Begin VB.Label num_label 
         Caption         =   "Number of [CODED]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   21
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label Label22 
         Caption         =   "Lenght"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3330
         TabIndex        =   20
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label21 
         Caption         =   "Width"
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
         Left            =   540
         TabIndex        =   19
         Top             =   1650
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Height"
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
         Left            =   480
         TabIndex        =   18
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Label Label19 
         Caption         =   "Cross-Section :"
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
         Left            =   180
         TabIndex        =   17
         Top             =   930
         Width           =   1425
      End
      Begin VB.Label Label9 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   210
         TabIndex        =   33
         Top             =   2580
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   39
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   3690
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1590
         TabIndex        =   40
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   41
         Top             =   2580
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3150
         TabIndex        =   34
         Top             =   2580
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3060
         TabIndex        =   38
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4410
         TabIndex        =   35
         Top             =   2580
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "f"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4410
         TabIndex        =   36
         Top             =   3150
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Concrete Composition"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   3765
      Begin VB.TextBox cement_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   1
         Top             =   330
         Width           =   705
      End
      Begin VB.TextBox aggregates_txt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   780
         Width           =   705
      End
      Begin VB.Label Label28 
         Caption         =   "Kg/m3"
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
         Left            =   2220
         TabIndex        =   47
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label27 
         Caption         =   "Kg/m3"
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
         Left            =   2220
         TabIndex        =   46
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Cement"
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
         Left            =   600
         TabIndex        =   45
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label23 
         Caption         =   "Aggregates"
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
         Left            =   330
         TabIndex        =   44
         Top             =   810
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   480
      Picture         =   "frm_concrete_board2.frx":324A
      Top             =   2160
      Width           =   3000
   End
End
Attribute VB_Name = "frm_concrete_board2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db_name As String
Public db_pos As Integer

Private Sub Close_Click()
    frm_concrete_board2.Hide
    Unload Me
End Sub

Private Sub Command1_Click()
Dim doc As Integer

If Not validate_fields() Then
    Exit Sub
End If
doc = current_form()
FState(doc).count = FState(doc).count + 1
With document(doc)
    If .lista.Rows <= FState(doc).count Then
        .lista.Rows = .lista.Rows + 10
    End If
    With .lista
        .Row = FState(doc).count
        .Col = 0
        If struct_type.Text = "Beams" Then
            .Text = "Conc.Beam"
        Else
            .Text = "Conc.Pillar"
        End If
        .Col = 1
        .Text = num_beams.Text
        .Col = 2
        .Text = height_txt.Text
        .Col = 3
        .Text = width_txt.Text
        .Col = 4
        .Text = "-"
        .Col = 5
        .Text = lenght_txt.Text
        .Col = 6
        .Text = f5_txt.Text
        .Col = 7
        .Text = f6_txt.Text
        .Col = 8
        .Text = f8_txt.Text
        .Col = 9
        .Text = f10_txt.Text
        .Col = 10
        .Text = f12_txt.Text
        .Col = 11
        .Text = f16_txt.Text
        .Col = 12
        .Text = f20_txt.Text
        .Col = 13
        .Text = f25_txt.Text
        .Col = 14
        .Text = f32_txt.Text
        .Col = 15
        .Text = cement_txt.Text
        .Col = 16
        .Text = aggregates_txt.Text
        .Col = 17
        .Text = costs_txt.Text
        .Col = 18
        .Text = Me.db_name
        .Col = 19
        .Text = Me.db_pos
    End With
End With
End Sub

Private Sub Command2_Click()
Unload Me
frm_concrete_board3.Show 1
End Sub


Private Sub Form_Load()

Call DisableX(frm_concrete_board2)
struct_type.AddItem "Beams"
struct_type.AddItem "Pillars"
struct_type.ListIndex = 0
Frame2.Caption = struct_type.Text

num_label.Caption = "number of " & struct_type.Text
If enabler("Concrete", "Costs") <> "Null" Then
    costs_txt.Text = enabler("Concrete", "Costs")
    costs_txt.Enabled = False
End If
If enabler("Concrete", "Aggregates") <> "Null" Then
    aggregates_txt.Text = enabler("Concrete", "Aggregates")
    aggregates_txt.Enabled = False
End If
If enabler("Concrete", "Cement") <> "Null" Then
    cement_txt.Text = enabler("Concrete", "Cement")
    cement_txt.Enabled = False
End If
End Sub

Private Function validate_fields() As Boolean

validate_fields = True
If Not IsNumeric(costs_txt.Text) Then
    validate_fields = False
    costs_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(aggregates_txt.Text) Then
    validate_fields = False
    aggregates_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(cement_txt.Text) Then
    validate_fields = False
    cement_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(height_txt.Text) Then
    validate_fields = False
    height_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(width_txt.Text) Then
    validate_fields = False
    width_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(lenght_txt.Text) Then
    validate_fields = False
    lenght_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(num_beams.Text) Then
    validate_fields = False
    num_beams.SetFocus
    Exit Function
End If
If Not IsNumeric(f5_txt.Text) Then
    validate_fields = False
    f5_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f6_txt.Text) Then
    validate_fields = False
    f6_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f8_txt.Text) Then
    validate_fields = False
    f8_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f10_txt.Text) Then
    validate_fields = False
    f10_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f12_txt.Text) Then
    validate_fields = False
    f12_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f16_txt.Text) Then
    validate_fields = False
    f16_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f20_txt.Text) Then
    validate_fields = False
    f20_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f25_txt.Text) Then
    validate_fields = False
    f25_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(f32_txt.Text) Then
    validate_fields = False
    f32_txt.SetFocus
    Exit Function
End If

End Function

Private Sub struct_type_Click()
Frame2.Caption = struct_type.Text
num_label.Caption = "number of " & struct_type.Text

End Sub
