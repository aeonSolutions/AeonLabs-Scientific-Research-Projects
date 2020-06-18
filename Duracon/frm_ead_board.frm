VERSION 5.00
Begin VB.Form frm_ead_board 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DURACON - General Information & Distribution Data"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "frm_ead_board.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_close 
      Caption         =   "Close"
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
      Left            =   4830
      TabIndex        =   43
      Top             =   7140
      Width           =   1425
   End
   Begin VB.Frame Frame4 
      Caption         =   "Geometric Parameters"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   150
      TabIndex        =   25
      Top             =   5820
      Width           =   7695
      Begin VB.ComboBox Distype 
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
         Index           =   0
         ItemData        =   "frm_ead_board.frx":324A
         Left            =   390
         List            =   "frm_ead_board.frx":325A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   570
         Width           =   1695
      End
      Begin VB.TextBox Param2 
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
         Index           =   0
         Left            =   5820
         TabIndex        =   17
         Top             =   570
         Width           =   1695
      End
      Begin VB.TextBox Param1 
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
         Index           =   0
         Left            =   3930
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000015&
         X1              =   3930
         X2              =   7500
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label20 
         Caption         =   "Parameter 1"
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
         Left            =   4230
         TabIndex        =   42
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "Parameters 2"
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
         Left            =   6090
         TabIndex        =   41
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Concrete Cover"
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
         Left            =   180
         TabIndex        =   29
         Top             =   330
         Width           =   1605
      End
      Begin VB.Label Label6 
         Caption         =   "Xc (mm)"
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
         Left            =   2460
         TabIndex        =   28
         Top             =   570
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Concrete Quality"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   150
      TabIndex        =   24
      Top             =   3180
      Width           =   7695
      Begin VB.TextBox Param2 
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
         Index           =   4
         Left            =   5790
         TabIndex        =   14
         Top             =   1770
         Width           =   1695
      End
      Begin VB.TextBox Param1 
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
         Index           =   4
         Left            =   3900
         TabIndex        =   13
         Top             =   1770
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
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
         Index           =   4
         ItemData        =   "frm_ead_board.frx":3286
         Left            =   360
         List            =   "frm_ead_board.frx":3296
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1770
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
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
         Index           =   1
         ItemData        =   "frm_ead_board.frx":32C2
         Left            =   360
         List            =   "frm_ead_board.frx":32D2
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox Param2 
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
         Index           =   1
         Left            =   5790
         TabIndex        =   8
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox Param2 
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
         Index           =   2
         Left            =   5790
         TabIndex        =   11
         Top             =   1170
         Width           =   1695
      End
      Begin VB.TextBox Param1 
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
         Index           =   1
         Left            =   3900
         TabIndex        =   7
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox Param1 
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
         Index           =   2
         Left            =   3900
         TabIndex        =   10
         Top             =   1170
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
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
         Index           =   2
         ItemData        =   "frm_ead_board.frx":32FE
         Left            =   360
         List            =   "frm_ead_board.frx":330E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000015&
         X1              =   3930
         X2              =   7500
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label17 
         Caption         =   "Parameter 1"
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
         Left            =   4230
         TabIndex        =   40
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Parameters 2"
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
         Left            =   6090
         TabIndex        =   39
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Age Effect Diffusion"
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
         Left            =   150
         TabIndex        =   36
         Top             =   1560
         Width           =   2235
      End
      Begin VB.Label Label19 
         Caption         =   "(-)"
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
         Left            =   2700
         TabIndex        =   35
         Top             =   1770
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2550
         TabIndex        =   34
         Top             =   1770
         Width           =   135
      End
      Begin VB.Label Label7 
         Caption         =   "Ccr (% wt./cem)"
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
         TabIndex        =   33
         Top             =   1170
         Width           =   1665
      End
      Begin VB.Label Label8 
         Caption         =   "Critical Chloride Concentration"
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
         Left            =   150
         TabIndex        =   32
         Top             =   930
         Width           =   3135
      End
      Begin VB.Label Label9 
         Caption         =   "Deff (e-12 m2/s)"
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
         Left            =   2190
         TabIndex        =   31
         Top             =   540
         Width           =   1635
      End
      Begin VB.Label Label13 
         Caption         =   "Diffusion Coefficient"
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
         Left            =   150
         TabIndex        =   30
         Top             =   300
         Width           =   2925
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Loading Conditions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   150
      TabIndex        =   23
      Top             =   1860
      Width           =   7695
      Begin VB.TextBox Param2 
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
         Index           =   3
         Left            =   5850
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Param1 
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
         Index           =   3
         Left            =   3960
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
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
         Index           =   3
         ItemData        =   "frm_ead_board.frx":333A
         Left            =   420
         List            =   "frm_ead_board.frx":334A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   3960
         X2              =   7530
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Label Label16 
         Caption         =   "Parameters 2"
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
         Left            =   6120
         TabIndex        =   38
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Parameter 1"
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
         Left            =   4260
         TabIndex        =   37
         Top             =   330
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Surface Chloride Concentration"
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
         Left            =   210
         TabIndex        =   27
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label11 
         Caption         =   "Cs (% wt./cem)"
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
         Left            =   2250
         TabIndex        =   26
         Top             =   660
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Project Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   150
      TabIndex        =   18
      Top             =   120
      Width           =   7695
      Begin VB.TextBox Prjname 
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
         Left            =   870
         TabIndex        =   0
         Top             =   360
         Width           =   6495
      End
      Begin VB.TextBox Datepj 
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
         Left            =   780
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Description 
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
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   6045
      End
      Begin VB.Label Label1 
         Caption         =   "Name :"
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
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Description :"
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
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Date :"
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
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   675
      End
   End
   Begin VB.CommandButton cmd_save 
      Caption         =   "Finish"
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
      Left            =   6360
      TabIndex        =   19
      Top             =   7140
      Width           =   1425
   End
End
Attribute VB_Name = "frm_ead_board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doc As Integer

Private Sub cmd_save_Click()
   Dim tmp() As String
    For i = 0 To 4
    With doc_props(doc)
        If Param1(i).Text <> "" Then
            tmp() = Split(Param1(i).Text, ".")
            If tmp(0) = "." Or tmp(0) = "" Then
                Param1(i).Text = "0" & Param1(i).Text
            End If
            tmp() = Split(Param1(i).Text, ",")
            If tmp(0) = "," Or tmp(0) = "" Then
                Param1(i).Text = "0" & Param1(i).Text
            End If
            If IsNumeric(Param1(i).Text) Then
                On Error Resume Next
                .prmvone(i) = Val(Param1(i).Text)
                If Err.Number <> 0 Then
                    If Err.Number = 6 Then
                        MsgBox "Numeric value too large", vbOKCancel + vbCritical, "duracon"
                        Param1(i).SetFocus
                        save_data = False
                        Exit Sub
                    Else
                        MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOKCancel + vbCritical, "duracon"
                        Param1(i).SetFocus
                        save_data = False
                        Exit Sub
                    End If
                End If
            End If
        Else
            MsgBox "Only Positive numeric Values allowed!", vbOKCancel + vbCritical, "duracon"
            Param1(i).SetFocus
            save_data = False
            Exit Sub
        End If
        If Param1(i).Text <> "" Then
            tmp() = Split(Param2(i).Text, ".")
            If tmp(0) = "." Or tmp(0) = "" Then
                Param2(i).Text = "0" & Param2(i).Text
            End If
            tmp() = Split(Param2(i).Text, ",")
            If tmp(0) = "," Or tmp(0) = "" Then
                Param2(i).Text = "0" & Param2(i).Text
            End If
            If IsNumeric(Param2(i).Text) Then
                On Error Resume Next
                .prmvtwo(i) = Val(Param2(i).Text)
                If Err.Number <> 0 Then
                    If Err.Number = 6 Then
                        MsgBox "Numeric value too large", vbOKCancel + vbCritical, "duracon"
                        Param2(i).SetFocus
                        save_data = False
                        Exit Sub
                    Else
                        MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOKCancel + vbCritical, "duracon"
                        Param2(i).SetFocus
                        save_data = False
                        Exit Sub
                    End If
                End If
            End If
        Else
            MsgBox "Only Positive numeric Values allowed!", vbOKCancel + vbCritical, "duracon"
            Param2(i).SetFocus
            save_data = False
            Exit Sub
        End If
    
    End With
    Next i
        
    With doc_props(doc)
        .frm_ead_board_values.project_name = Prjname.Text
        .nprojt = Prjname.Text
        .descrip = Description.Text
        .datepjt = Datepj.Text
        .frm_ead_board_values.Description = Description.Text
        .frm_ead_board_values.project_date = Datepj.Text
        .frm_ead_board_values.values = True
        For i = 0 To 4
            If Distype(i).Text = "Deterministic" Then .prmdistn(i) = 7
            If Distype(i).Text = "Normal" Then .prmdistn(i) = 0
            If Distype(i).Text = "Lognormal" Then .prmdistn(i) = 1
            If Distype(i).Text = "Beta" Then .prmdistn(i) = 2
            .frm_ead_board_values.Distype(i) = Distype(i).ListIndex
        Next i
        .prmdistn(5) = 7
        .prmdistn(6) = 7
        .prmdistn(7) = 0
        .frm_ead_board_values.ready = True
    End With
    Call refresh_lista(doc)
    FState(doc).values = True
    FState(doc).saved = False
    frm_ead_board.Hide
    Unload Me
End Sub

Private Sub cmd_close_Click()
    doc_props(doc).frm_ead_board_values.values = False
    FState(doc).values = False
    Call refresh_lista(doc)
    frm_ead_board.Hide
    Unload Me
End Sub

Private Sub Form_Load()
Call DisableX(frm_ead_board)
doc = current_form()
With doc_props(doc)
    'flaging the form
    .frm_ca_board1_values.cdc = -1
    
    .iseedv = 1
    .nsimul = 10000
    .seed = 0
    ' Desgin value default values
    .idifcoef = 1
    .kk = 123
    ' average anual temperature
    .prmvone(7) = 21
    .prmvtwo(7) = 0
    'age of structure during assessment
    .prmvone(6) = 28
    .prmvtwo(6) = 0
    'design life of structure - default is 10 years
    .tseriev = 10
    .prmvone(5) = 0
    .prmvtwo(5) = 50
    If .frm_ead_board_values.values And .frm_ead_board_values.project_name <> "" Then ' there's already input data stored
        With .frm_ead_board_values
            Prjname.Text = .project_name
            Description.Text = .Description
            Datepj.Text = .project_date
        End With
        For i = 0 To 4
            Distype(i).ListIndex = .frm_ead_board_values.Distype(i)
            Param1(i).Text = CStr(.prmvone(i))
            Param2(i).Text = CStr(.prmvtwo(i))
        Next i
    Else
        For i = 0 To 4
            Distype(i).ListIndex = 1
        Next i
    End If
    If .datepjt = "" Then
        .datepjt = Date
        Datepj.Text = Date
    Else
        Datepj.Text = .datepjt
    End If
End With
End Sub
