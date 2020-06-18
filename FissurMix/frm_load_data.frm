VERSION 5.00
Begin VB.Form frm_load_data 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[CODED]"
   ClientHeight    =   4275
   ClientLeft      =   3060
   ClientTop       =   2820
   ClientWidth     =   7485
   Icon            =   "frm_load_data.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Next_button 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5970
      TabIndex        =   3
      Top             =   3720
      Width           =   1365
   End
   Begin VB.CommandButton Cancel_button 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4470
      TabIndex        =   2
      Top             =   3720
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Experimental Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7215
      Begin VB.ComboBox scale_Combo 
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
         ItemData        =   "frm_load_data.frx":08CA
         Left            =   3270
         List            =   "frm_load_data.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2490
         Width           =   1395
      End
      Begin VB.TextBox emax_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   4470
         TabIndex        =   22
         Text            =   "0,9"
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox n_points_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   2430
         TabIndex        =   21
         Text            =   "0,05"
         Top             =   1890
         Width           =   825
      End
      Begin VB.TextBox c_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   6120
         TabIndex        =   18
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox x_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   4980
         TabIndex        =   17
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox x2_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   3780
         TabIndex        =   16
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox x3_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   2580
         TabIndex        =   15
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox x4_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   1380
         TabIndex        =   14
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox x5_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   1
         EndProperty
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
         Left            =   180
         TabIndex        =   13
         Top             =   1260
         Width           =   825
      End
      Begin VB.Label Label17 
         Caption         =   "%"
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
         Left            =   3300
         TabIndex        =   31
         Top             =   1950
         Width           =   465
      End
      Begin VB.Label Label16 
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   30
         Top             =   1830
         Width           =   165
      End
      Begin VB.Label Label11 
         Caption         =   "max"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4080
         TabIndex        =   20
         Top             =   1950
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   1830
         Width           =   165
      End
      Begin VB.Label Label15 
         Caption         =   "Scale Type"
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
         Left            =   2100
         TabIndex        =   24
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label scale_info_1 
         Caption         =   "%"
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
         Left            =   5340
         TabIndex        =   23
         Top             =   1920
         Width           =   465
      End
      Begin VB.Label Label10 
         Caption         =   "intervals of"
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
         Left            =   1200
         TabIndex        =   19
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "2"
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
         Left            =   4710
         TabIndex        =   12
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label Label8 
         Caption         =   "3"
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
         Left            =   3510
         TabIndex        =   11
         Top             =   1230
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "4"
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
         Left            =   2340
         TabIndex        =   10
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   1110
         TabIndex        =   9
         Top             =   1230
         Width           =   105
      End
      Begin VB.Label Label5 
         Caption         =   "x +"
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
         Left            =   5820
         TabIndex        =   8
         Top             =   1290
         Width           =   345
      End
      Begin VB.Label Label4 
         Caption         =   "x  +"
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
         Left            =   4620
         TabIndex        =   7
         Top             =   1290
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "x  +"
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
         Left            =   3420
         TabIndex        =   6
         Top             =   1290
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "x  +"
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
         Left            =   2220
         TabIndex        =   5
         Top             =   1290
         Width           =   405
      End
      Begin VB.Label num_concrete_entries 
         Caption         =   "Please insert the terms of a 5 degree polinomial function that best fits your experimental data:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   300
         TabIndex        =   4
         Top             =   480
         Width           =   6645
      End
      Begin VB.Label Label1 
         Caption         =   "x  +"
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
         Left            =   1020
         TabIndex        =   1
         Top             =   1290
         Width           =   405
      End
      Begin VB.Label Label14 
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3180
         TabIndex        =   28
         Top             =   870
         Width           =   135
      End
      Begin VB.Label Label13 
         Caption         =   "F(  ), "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2970
         TabIndex        =   27
         Top             =   930
         Width           =   495
      End
      Begin VB.Label scale_txt 
         Caption         =   "[CODED]"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3510
         TabIndex        =   29
         Top             =   930
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frm_load_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private doc As Integer

Private Sub Cancel_button_Click()
    Call reset_props
    Call refresh_richtext
    Me.Hide
    Unload Me
End Sub

Private Sub scale_combo_Click()
    With doc_props(doc)
        With .exp_data
            .scales = scale_Combo.Text
            scale_txt.Caption = " [" & .scales & "/ %] "
        End With
    End With
End Sub

Private Sub Form_Load()
    doc = current_form
    Call DisableX(frm_load_data)
    Me.Caption = App.Title
    If doc_props(doc).exp_data.scales <> "-1" And doc_props(doc).exp_data.scales <> "" Then
        With doc_props(doc).exp_data
            x5_txt.Text = .x5
            x4_txt.Text = .x4
            x3_txt.Text = .x3
            x2_txt.Text = .x2
            x_txt.Text = .x
            c_txt.Text = .c
            emax_txt.Text = .emax * 100
            n_points_txt.Text = .delta_e * 100
            scale_Combo.Clear
            scale_Combo.AddItem "MN"
            scale_Combo.AddItem "KN"
            scale_Combo.AddItem "N"
            If .scales = "MN" Then
                scale_Combo.ListIndex = 0
            ElseIf .scales = "KN" Then
                scale_Combo.ListIndex = 1
            ElseIf .scales = "N" Then
                scale_Combo.ListIndex = 2
            End If
        End With
    Else
        scale_Combo.Clear
        scale_Combo.AddItem "MN"
        scale_Combo.AddItem "KN"
        scale_Combo.AddItem "N"
        scale_Combo.ListIndex = 1
    End If
End Sub
Private Sub Next_button_Click()
    If Not IsNumeric(x5_txt.Text) Then
        x5_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(x4_txt.Text) Then
        x4_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(x3_txt.Text) Then
        x3_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(x2_txt.Text) Then
        x2_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(x_txt.Text) Then
        x_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(c_txt.Text) Then
        c_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(n_points_txt.Text) Then
        n_points_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(emax_txt.Text) Then
        dmax_txt.SetFocus
        Exit Sub
    End If
    
    With doc_props(doc)
        With .exp_data
            .scales = scale_Combo.Text
            .x5 = convert_type(x5_txt.Text)
            .x4 = convert_type(x4_txt.Text)
            .x3 = convert_type(x3_txt.Text)
            .x2 = convert_type(x2_txt.Text)
            .x = convert_type(x_txt.Text)
            .c = convert_type(c_txt.Text)
            .emax = convert_type(emax_txt.Text) / 100 ' converted from % to adimensional
            .delta_e = convert_type(n_points_txt.Text) / 100 ' converted from % to adimensional
        End With
    End With
    
   
    Me.Hide
    Unload Me
    Call refresh_richtext
    frm_load_data2.Show 1
End Sub

