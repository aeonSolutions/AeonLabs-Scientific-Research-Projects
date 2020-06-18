VERSION 5.00
Begin VB.Form frm_concrete_board3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pillars"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6585
   Icon            =   "frm_concrete_board3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   345
      Left            =   2550
      TabIndex        =   40
      Top             =   4650
      Width           =   1125
   End
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   345
      Left            =   5190
      TabIndex        =   39
      Top             =   4650
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   345
      Left            =   3870
      TabIndex        =   38
      Top             =   4650
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pillars"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   300
      TabIndex        =   32
      Top             =   390
      Width           =   6015
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
         Height          =   345
         Left            =   4830
         TabIndex        =   4
         Top             =   960
         Width           =   705
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
         Height          =   345
         Left            =   4830
         TabIndex        =   3
         Top             =   570
         Width           =   705
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
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   990
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
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   570
         Width           =   705
      End
      Begin VB.Label Label23 
         Caption         =   "Number of Pillars"
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
         Left            =   3210
         TabIndex        =   37
         Top             =   1020
         Width           =   1605
      End
      Begin VB.Label Label22 
         Caption         =   "Pillar Lenght (m)"
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
         Left            =   3210
         TabIndex        =   36
         Top             =   630
         Width           =   1605
      End
      Begin VB.Label Label21 
         Caption         =   "Width (m)"
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
         Left            =   870
         TabIndex        =   35
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label20 
         Caption         =   "Height (m)"
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
         Left            =   840
         TabIndex        =   34
         Top             =   600
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
         Left            =   330
         TabIndex        =   33
         Top             =   330
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Steel bars"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   300
      TabIndex        =   0
      Top             =   2370
      Width           =   6015
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
         Height          =   315
         Left            =   690
         TabIndex        =   13
         Text            =   "0"
         Top             =   1440
         Width           =   705
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
         Height          =   315
         Left            =   5010
         TabIndex        =   12
         Text            =   "0"
         Top             =   900
         Width           =   705
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
         Height          =   315
         Left            =   3630
         TabIndex        =   11
         Text            =   "0"
         Top             =   900
         Width           =   705
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
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Text            =   "0"
         Top             =   900
         Width           =   705
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
         Height          =   315
         Left            =   690
         TabIndex        =   9
         Text            =   "0"
         Top             =   900
         Width           =   705
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
         Height          =   315
         Left            =   5010
         TabIndex        =   8
         Text            =   "0"
         Top             =   360
         Width           =   705
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
         Height          =   315
         Left            =   3630
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   705
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
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Text            =   "0"
         Top             =   360
         Width           =   705
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
         Height          =   315
         Left            =   690
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   705
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
         Left            =   420
         TabIndex        =   31
         Top             =   1500
         Width           =   345
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
         Left            =   4710
         TabIndex        =   30
         Top             =   960
         Width           =   345
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
         Left            =   3360
         TabIndex        =   29
         Top             =   930
         Width           =   345
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
         Left            =   1890
         TabIndex        =   28
         Top             =   930
         Width           =   345
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
         Left            =   450
         TabIndex        =   27
         Top             =   930
         Width           =   345
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
         Left            =   4710
         TabIndex        =   26
         Top             =   390
         Width           =   345
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
         Left            =   3450
         TabIndex        =   25
         Top             =   390
         Width           =   345
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
         Left            =   1980
         TabIndex        =   24
         Top             =   390
         Width           =   345
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
         Left            =   510
         TabIndex        =   23
         Top             =   390
         Width           =   345
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
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   225
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
         Left            =   3300
         TabIndex        =   21
         Top             =   360
         Width           =   225
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
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   225
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
         Left            =   4560
         TabIndex        =   19
         Top             =   930
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
         Left            =   270
         TabIndex        =   18
         Top             =   1470
         Width           =   225
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
         Left            =   3210
         TabIndex        =   17
         Top             =   900
         Width           =   225
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
         Left            =   300
         TabIndex        =   16
         Top             =   900
         Width           =   225
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
         Left            =   1740
         TabIndex        =   15
         Top             =   900
         Width           =   225
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
         Left            =   1830
         TabIndex        =   14
         Top             =   360
         Width           =   225
      End
   End
End
Attribute VB_Name = "frm_concrete_board3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
    frm_concrete_board3.Hide
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
    With .lista
        .Row = FState(doc).count
        .Col = 0
        .Text = "Conc.Pillar"
        .Col = 1
        .Text = num_beams.Text
        .Col = 2
        .Text = height_txt.Text
        .Col = 3
        .Text = width_txt.Text
        .Col = 4
        .Text = lenght_txt.Text
        .Col = 5
        .Text = f5_txt.Text
        .Col = 6
        .Text = f6_txt.Text
        .Col = 7
        .Text = f8_txt.Text
        .Col = 8
        .Text = f10_txt.Text
        .Col = 9
        .Text = f12_txt.Text
        .Col = 10
        .Text = f16_txt.Text
        .Col = 11
        .Text = f20_txt.Text
        .Col = 12
        .Text = f25_txt.Text
        .Col = 13
        .Text = f32_txt.Text
    End With
End With
End Sub

Private Sub Command2_Click()
Unload Me
frm_concrete_board2.Show 1
End Sub

Private Sub Form_Load()
Call DisableX(frm_concrete_board3)
End Sub

Private Function validate_fields() As Boolean

validate_fields = True
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
