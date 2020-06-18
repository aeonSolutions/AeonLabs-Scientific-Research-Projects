VERSION 5.00
Begin VB.Form frm_extension_selection 
   Caption         =   "Extension Selection"
   ClientHeight    =   1770
   ClientLeft      =   3555
   ClientTop       =   3405
   ClientWidth     =   6210
   Icon            =   "frm_extension_selection.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   150
      Top             =   1260
   End
   Begin VB.TextBox res_txt 
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
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   540
      Width           =   825
   End
   Begin VB.TextBox delta_e_txt 
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
      Left            =   2460
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   540
      Width           =   825
   End
   Begin VB.TextBox int_txt 
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
      Left            =   1470
      TabIndex        =   5
      Top             =   540
      Width           =   825
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
      Left            =   4740
      TabIndex        =   1
      Top             =   1290
      Width           =   1365
   End
   Begin VB.CommandButton ok_button 
      Caption         =   "ok"
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
      Left            =   3300
      TabIndex        =   0
      Top             =   1290
      Width           =   1365
   End
   Begin VB.Label info_txt 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   60
      TabIndex        =   11
      Top             =   900
      Width           =   6105
   End
   Begin VB.Label Label4 
      Caption         =   ":"
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
      Left            =   1350
      TabIndex        =   10
      Top             =   570
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "="
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
      Left            =   3330
      TabIndex        =   8
      Top             =   540
      Width           =   165
   End
   Begin VB.Label Label2 
      Caption         =   "x"
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
      Left            =   2310
      TabIndex        =   6
      Top             =   570
      Width           =   135
   End
   Begin VB.Label Label24 
      Caption         =   "(%)"
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
      Left            =   4350
      TabIndex        =   4
      Top             =   570
      Width           =   405
   End
   Begin VB.Label Label26 
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
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   510
      Width           =   165
   End
   Begin VB.Label label_txt 
      Alignment       =   2  'Center
      Caption         =   "Please insert a positive integer number not larger than"
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
      Left            =   90
      TabIndex        =   2
      Top             =   150
      Width           =   6075
   End
End
Attribute VB_Name = "frm_extension_selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim doc As Integer

Private Sub Cancel_button_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
Call DisableX(frm_extension_selection)

doc = current_form
With doc_props(doc)
    delta_e_txt.Text = .exp_data.delta_e * 100
    int_txt.Text = .results.live_data_pos
    label_txt.Caption = label_txt.Caption & Str(Round(.exp_data.emax / .exp_data.delta_e, 0)) & "."
End With
End Sub

Private Sub ok_button_Click()
Dim doc As Integer

doc = current_form
With doc_props(doc)
    If Not (IsNumeric(res_txt.Text)) Then
        res_txt.SetFocus
        Exit Sub
    End If
    If Round(.exp_data.emax / .exp_data.delta_e, 0) < convert_type(res_txt.Text) Then
        res_txt.SetFocus
        Exit Sub
    End If
    .results.live_data_pos = convert_type(int_txt.Text)
    
End With
Me.Hide
Unload Me
End Sub

Private Sub Timer1_Timer()
With doc_props(doc)
    res_txt.Text = convert_type(delta_e_txt.Text) * convert_type(int_txt.Text)
    If Not (IsNumeric(int_txt.Text)) Then
        info_txt.Caption = "Non Numerical value entered!"
    ElseIf Round(.exp_data.emax / .exp_data.delta_e, 0) < convert_type(int_txt.Text) Then
        info_txt.Caption = "Invalid value size.Please input a smaller value."
    ElseIf info_txt.Caption <> "" Then
        info_txt.Caption = ""
    End If
End With
End Sub
