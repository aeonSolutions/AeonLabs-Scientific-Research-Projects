VERSION 5.00
Begin VB.Form frm_edit_material 
   Caption         =   "Edit material"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   Icon            =   "frm_edit_material.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "Apply"
      Height          =   300
      Left            =   4620
      TabIndex        =   26
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   300
      Left            =   1950
      TabIndex        =   18
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3270
      TabIndex        =   17
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Material properties"
      ForeColor       =   &H80000008&
      Height          =   3960
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5655
      Begin VB.TextBox e_txt 
         Height          =   285
         Left            =   4110
         TabIndex        =   25
         Top             =   2070
         Width           =   1065
      End
      Begin VB.TextBox alfa_txt 
         Height          =   285
         Left            =   1410
         TabIndex        =   23
         Top             =   2880
         Width           =   1065
      End
      Begin VB.ComboBox combo 
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
         Left            =   2070
         TabIndex        =   19
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox g0_txt 
         Height          =   285
         Left            =   1410
         TabIndex        =   7
         Top             =   2490
         Width           =   1065
      End
      Begin VB.TextBox t1_txt 
         Height          =   285
         Left            =   1410
         TabIndex        =   6
         Top             =   2070
         Width           =   1065
      End
      Begin VB.TextBox area_txt 
         Height          =   285
         Left            =   4110
         TabIndex        =   5
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox n_txt 
         Height          =   285
         Left            =   4110
         TabIndex        =   4
         Top             =   2490
         Width           =   1065
      End
      Begin VB.TextBox q0_txt 
         Height          =   285
         Left            =   1395
         TabIndex        =   3
         Top             =   1260
         Width           =   1065
      End
      Begin VB.TextBox k_txt 
         Height          =   285
         Left            =   1410
         TabIndex        =   2
         Top             =   1680
         Width           =   1065
      End
      Begin VB.TextBox l_txt 
         Height          =   285
         Left            =   4110
         TabIndex        =   1
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "-5"
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
         Left            =   1140
         TabIndex        =   27
         Top             =   2790
         Width           =   315
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   2820
         X2              =   2820
         Y1              =   1170
         Y2              =   3660
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   2850
         X2              =   2850
         Y1              =   1170
         Y2              =   3660
      End
      Begin VB.Label Label11 
         Caption         =   "E (GPa)"
         Height          =   255
         Left            =   3420
         TabIndex        =   24
         Top             =   2100
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "/ºC x10"
         Height          =   225
         Left            =   570
         TabIndex        =   22
         Top             =   2880
         Width           =   645
      End
      Begin VB.Label Label8 
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   21
         Top             =   2850
         Width           =   165
      End
      Begin VB.Label Label7 
         Caption         =   "Edit:"
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
         Left            =   1590
         TabIndex        =   20
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label18 
         Caption         =   "G0 (W/m2)"
         Height          =   225
         Left            =   390
         TabIndex        =   16
         Top             =   2550
         Width           =   1005
      End
      Begin VB.Label Label17 
         Caption         =   "T1 (ºC)"
         Height          =   225
         Left            =   750
         TabIndex        =   15
         Top             =   2130
         Width           =   585
      End
      Begin VB.Label Label9 
         Caption         =   "Area (mm2)"
         Height          =   225
         Left            =   3150
         TabIndex        =   14
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label4 
         Caption         =   "Split"
         Height          =   255
         Left            =   3450
         TabIndex        =   13
         Top             =   2520
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "q0 (W/mm3)"
         Height          =   270
         Left            =   405
         TabIndex        =   12
         Top             =   1260
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "k (W/m.ºC)"
         Height          =   270
         Left            =   450
         TabIndex        =   11
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "L (mm)"
         Height          =   225
         Left            =   3480
         TabIndex        =   10
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "Phisical Properties:"
         Height          =   225
         Left            =   3150
         TabIndex        =   9
         Top             =   900
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Thermal properties:"
         Height          =   315
         Left            =   150
         TabIndex        =   8
         Top             =   900
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frm_edit_material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lastindex As Integer


Private Sub CmdApply_Click()
   Dim i As Integer
   Dim j As Integer
   
   
   If Not IsNumeric(l_txt) Then
        l_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(area_txt) Then
        area_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(t1_txt) Then
        t1_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(g0_txt) Then
        g0_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(k_txt) Then
        k_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(q0_txt) Then
        q0_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(n_txt) Then
        n_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(e_txt) Then
        e_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(alfa_txt) Then
        alfa_txt.SetFocus
        Exit Sub
    End If
    'get the current active form
    i = current_form()
    FState(i).saved = False
    With document(i)
    .lista.col = 1
    .lista.row = lastindex + 1
    .lista.col = 1
    .lista.CellAlignment = 4
    .lista.Text = l_txt.Text
    .lista.col = 3
    .lista.CellAlignment = 4
    .lista.Text = k_txt.Text
    .lista.col = 7
    .lista.CellAlignment = 4
    .lista.Text = n_txt.Text
    .lista.col = 8
    .lista.CellAlignment = 4
    .lista.Text = e_txt.Text
    .lista.col = 9
    .lista.CellAlignment = 4
    .lista.Text = alfa_txt.Text
    For j = 1 To FState(i).Conta
        .lista.row = j
        .lista.col = 2
        .lista.CellAlignment = 4
        .lista.Text = area_txt.Text
        .lista.col = 4
        .lista.CellAlignment = 4
        .lista.Text = q0_txt.Text
        .lista.col = 5
        .lista.CellAlignment = 4
        .lista.Text = t1_txt.Text
        .lista.col = 6
        .lista.CellAlignment = 4
        .lista.Text = g0_txt.Text
    Next j
    FState(i).calculated = False
    .SSTab.TabEnabled(1) = False
    .SSTab.TabEnabled(2) = False
    .SSTab.TabEnabled(3) = False
    .lista.Refresh
    
End With

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call CmdApply_Click
    Unload Me

End Sub

Private Sub combo_Click()
If combo.ListIndex <> lastindex Then
    i = current_form
    lastindex = combo.ListIndex
    With document(i)
        With .lista
            .row = lastindex + 1
            .col = 1
            l_txt = .Text
            .col = 2
            area_txt = .Text
            .col = 3
            k_txt = .Text
            .col = 4
            q0_txt = .Text
            .col = 5
            t1_txt = .Text
            .col = 6
            g0_txt = .Text
            .col = 7
            n_txt.Text = .Text
            .col = 8
            e_txt.Text = .Text
            .col = 9
            alfa_txt.Text = .Text

        End With
    End With
End If
End Sub


Private Sub form_load()
    Dim i As Integer, j As Integer
    
    i = current_form
    Call DisableX(frm_edit_material)
    frm_edit_material.Caption = frm_edit_material.Caption & " - " & document(i).Caption
    'Dim activeChild As Form = Me.ActiveMDIChild
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    For j = 1 To FState(i).Conta - 1
        combo.AddItem "Material nº" & CStr(j)
    Next j
    combo.ListIndex = 0
    lastindex = 0
        With document(i)
            With .lista
                .row = 1
                .col = 1
                l_txt.Text = .Text
                .col = 2
                area_txt.Text = .Text
                .col = 3
                k_txt.Text = .Text
                .col = 4
                q0_txt.Text = .Text
                .col = 5
                t1_txt.Text = .Text
                .col = 6
                g0_txt.Text = .Text
                .col = 7
                n_txt.Text = .Text
                .col = 8
                e_txt.Text = .Text
                .col = 9
                alfa_txt.Text = .Text
            End With
        End With

End Sub

