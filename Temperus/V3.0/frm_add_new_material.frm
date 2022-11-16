VERSION 5.00
Begin VB.Form frm_add_new_material 
   Caption         =   "material properties"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   Icon            =   "frm_add_new_material.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "Add"
      Height          =   330
      Left            =   8610
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmd_delete 
      Caption         =   "Delete"
      Height          =   330
      Left            =   7320
      TabIndex        =   13
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   330
      Left            =   9900
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Material properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4560
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   10905
      Begin VB.ComboBox e_units 
         Enabled         =   0   'False
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
         Left            =   9030
         TabIndex        =   39
         Top             =   2550
         Width           =   1400
      End
      Begin VB.TextBox q0_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3210
         TabIndex        =   5
         Top             =   3480
         Width           =   1065
      End
      Begin VB.ComboBox q0_units 
         Enabled         =   0   'False
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
         Left            =   4440
         TabIndex        =   37
         Top             =   3480
         Width           =   1400
      End
      Begin VB.ComboBox mat_combo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4980
         TabIndex        =   36
         Top             =   450
         Width           =   2565
      End
      Begin VB.ComboBox l_units 
         Enabled         =   0   'False
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
         Left            =   9030
         TabIndex        =   34
         Top             =   1650
         Width           =   1400
      End
      Begin VB.ComboBox area_units 
         Enabled         =   0   'False
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
         Left            =   9030
         TabIndex        =   33
         Top             =   2130
         Width           =   1400
      End
      Begin VB.ComboBox k_units 
         Enabled         =   0   'False
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
         Left            =   4440
         TabIndex        =   32
         Top             =   2070
         Width           =   1400
      End
      Begin VB.ComboBox td_units 
         Enabled         =   0   'False
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
         Left            =   4440
         TabIndex        =   31
         Top             =   3000
         Width           =   1400
      End
      Begin VB.ComboBox te_units 
         Enabled         =   0   'False
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
         Left            =   4440
         TabIndex        =   30
         Top             =   2490
         Width           =   1400
      End
      Begin VB.ComboBox b_units 
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
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1590
         Width           =   1400
      End
      Begin VB.TextBox e_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   9
         Top             =   2580
         Width           =   1065
      End
      Begin VB.TextBox alfa_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3210
         TabIndex        =   6
         Top             =   3960
         Width           =   1065
      End
      Begin VB.TextBox td_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3210
         TabIndex        =   4
         Top             =   3000
         Width           =   1065
      End
      Begin VB.TextBox te_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3210
         TabIndex        =   3
         Top             =   2520
         Width           =   1065
      End
      Begin VB.TextBox area_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   8
         Top             =   2130
         Width           =   1065
      End
      Begin VB.TextBox n_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   10
         Top             =   3060
         Width           =   1065
      End
      Begin VB.TextBox b_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3195
         TabIndex        =   1
         Top             =   1590
         Width           =   1065
      End
      Begin VB.TextBox k_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3210
         TabIndex        =   2
         Top             =   2070
         Width           =   1065
      End
      Begin VB.TextBox l_txt 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7860
         TabIndex        =   7
         Top             =   1650
         Width           =   1065
      End
      Begin VB.Label Label14 
         Caption         =   "Q0 :"
         Height          =   225
         Left            =   2730
         TabIndex        =   38
         Top             =   3510
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   1950
         Left            =   210
         Picture         =   "frm_add_new_material.frx":08CA
         Top             =   1800
         Width           =   2100
      End
      Begin VB.Label tipo_txt 
         Caption         =   "Inserting material : "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   35
         Top             =   450
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Units:"
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
         Left            =   9120
         TabIndex        =   29
         Top             =   1290
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "Units:"
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
         Left            =   4500
         TabIndex        =   28
         Top             =   1230
         Width           =   1125
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
         Left            =   5010
         TabIndex        =   26
         Top             =   3930
         Width           =   315
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   6330
         X2              =   6330
         Y1              =   1230
         Y2              =   4230
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   6360
         X2              =   6360
         Y1              =   1230
         Y2              =   4170
      End
      Begin VB.Label Label11 
         Caption         =   "E :"
         Height          =   255
         Left            =   7440
         TabIndex        =   25
         Top             =   2610
         Width           =   315
      End
      Begin VB.Label Label10 
         Caption         =   "/ºC x10"
         Height          =   225
         Left            =   4470
         TabIndex        =   24
         Top             =   4020
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "a :"
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
         Left            =   2790
         TabIndex        =   23
         Top             =   3960
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Td :"
         Height          =   225
         Left            =   2730
         TabIndex        =   22
         Top             =   3030
         Width           =   405
      End
      Begin VB.Label Label17 
         Caption         =   "Te :"
         Height          =   225
         Left            =   2730
         TabIndex        =   21
         Top             =   2580
         Width           =   405
      End
      Begin VB.Label Label9 
         Caption         =   "Area :"
         Height          =   225
         Left            =   7230
         TabIndex        =   20
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "nº of elements:"
         Height          =   255
         Left            =   6660
         TabIndex        =   19
         Top             =   3090
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "b :"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2820
         TabIndex        =   18
         Top             =   1590
         Width           =   270
      End
      Begin VB.Label Label2 
         Caption         =   "k :"
         Height          =   270
         Left            =   2820
         TabIndex        =   17
         Top             =   2070
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "L :"
         Height          =   225
         Left            =   7440
         TabIndex        =   16
         Top             =   1650
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Phisical Properties:"
         Height          =   225
         Left            =   6900
         TabIndex        =   15
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Thermal properties:"
         Height          =   315
         Left            =   1950
         TabIndex        =   14
         Top             =   1230
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frm_add_new_material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim last_index As Integer
Dim apply As Boolean
Dim tipo As Boolean
Dim i As Integer


Public doc As Integer

Private Sub cmd_delete_Click()
Dim j As Integer

    If FState(doc).Conta = 1 Then
       MsgBox "There are no material to remove in the current document.", vbOK + vbCritical, " Temperus "
       Exit Sub
    End If
FState(doc).calculated = False
FState(doc).Conta = FState(doc).Conta - 1
With document(doc)
   .lista.RemoveItem (last_index + 1)
   .lista.AddItem ""
   .lista.col = 0
   If FState(doc).Conta > 1 Then
     For j = 1 To FState(doc).Conta - 1
       .lista.CellAlignment = 4
       .lista.CellFontBold = True
       .lista.row = j
       .lista.Text = Str(j)
     Next j
   End If
   .lista.Refresh
End With
mat_combo.Clear
For j = 1 To FState(doc).Conta
    mat_combo.AddItem "Material nº" & CStr(j)
Next j
mat_combo.ListIndex = FState(doc).Conta - 1
End Sub

Private Sub CmdApply_Click()
   Dim i As Integer
   Dim j As Integer
   Dim txt(10) As String
   
   If Not IsNumeric(l_txt) Then
        l_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(q0_txt) Then
        q0_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(area_txt) Then
        area_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(te_txt) Then
        te_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(b_txt) Then
        b_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(k_txt) Then
        k_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(td_txt) Then
        td_txt.SetFocus
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
    txt(1) = l_txt.Text
    txt(2) = k_txt.Text
    txt(3) = b_txt.Text
    txt(4) = te_txt.Text
    txt(5) = td_txt.Text
    txt(6) = n_txt.Text
    txt(7) = e_txt.Text
    txt(8) = alfa_txt.Text
    txt(9) = area_txt.Text
    txt(10) = q0_txt.Text

    
    If tipo Then ' new material
       i = FState(doc).Conta
       FState(doc).Conta = FState(doc).Conta + 1
       mat_combo.Clear
       For j = 1 To FState(doc).Conta
         mat_combo.AddItem "Material nº" & CStr(j)
       Next j
       mat_combo.ListIndex = FState(doc).Conta - 1
    Else
       i = last_index + 1
    End If

    FState(doc).saved = False
    With document(doc)
    If .lista.Rows <= i Then
        .lista.Rows = i + 1
    End If
    .lista.row = i
    .lista.col = 0
    .lista.CellAlignment = 4
    .lista.Text = Str(i)
    .lista.col = 1
    .lista.CellAlignment = 4
    .lista.Text = txt(1)
    .lista.col = 3
    .lista.CellAlignment = 4
    .lista.Text = txt(2)
    .lista.col = 7
    .lista.CellAlignment = 4
    .lista.Text = txt(6)
    .lista.col = 8
    .lista.CellAlignment = 4
    .lista.Text = txt(7)
    .lista.col = 9
    .lista.CellAlignment = 4
    .lista.Text = txt(8)
    .lista.col = 2
    .lista.CellAlignment = 4
    .lista.Text = txt(9)
    .lista.col = 4
    .lista.CellAlignment = 4
    .lista.Text = txt(3)
    .lista.col = 10
    .lista.CellAlignment = 4
    .lista.Text = txt(10)
    For j = 1 To i
        .lista.row = j
        .lista.col = 5
        .lista.CellAlignment = 4
        .lista.Text = txt(4)
        .lista.col = 6
        .lista.CellAlignment = 4
        .lista.Text = txt(5)
    Next j
    FState(doc).calculated = False
    .SSTab.TabEnabled(1) = False
    .SSTab.TabEnabled(2) = False
    .SSTab.TabEnabled(3) = False
    .lista.Refresh
    
End With
apply = True
End Sub

Private Sub cmdCancel_Click()
    If Not apply Then
        Call CmdApply_Click
    End If
    Unload Me
End Sub



Private Sub mat_combo_Click()
If mat_combo.ListIndex <> last_index Then
    On Error Resume Next
    i = current_form
    last_index = mat_combo.ListIndex
    With document(i)
        With .lista
            .row = last_index + 1
            .col = 1
            l_txt = .Text
            .col = 2
            area_txt = .Text
            .col = 3
            k_txt = .Text
            .col = 4
            b_txt = .Text
            .col = 5
            te_txt = .Text
            .col = 6
            td_txt = .Text
            .col = 7
            n_txt.Text = .Text
            .col = 8
            e_txt.Text = .Text
            .col = 9
            alfa_txt.Text = .Text
            .col = 10
            q0_txt.Text = .Text
        End With
    End With
End If
If mat_combo.ListCount = mat_combo.ListIndex + 1 Then
    CmdApply.Caption = "Add"
    tipo = True
    tipo_txt = "Inserting material : "
    cmd_delete.Enabled = False
Else
    CmdApply.Caption = "Edit"
    tipo = False
    tipo_txt = "Editting material : "
    cmd_delete.Enabled = True
End If

End Sub

Private Sub form_load()
    Dim tmp As String
    
    doc = current_form
    mat_combo.Clear
    For j = 1 To FState(doc).Conta
         mat_combo.AddItem "Material nº" & CStr(j)
    Next j
    mat_combo.ListIndex = FState(doc).Conta - 1
    mat_combo.Enabled = True
    tipo_txt = "Inserting material : "
    
    
    Call add_units_item(l_units, units.l.txt, units.l.selected - 1)
    Call add_units_item(area_units, units.area.txt, units.area.selected - 1)
    Call add_units_item(te_units, units.te.txt, units.te.selected - 1)
    Call add_units_item(td_units, units.td.txt, units.td.selected - 1)
    Call add_units_item(k_units, units.k.txt, units.k.selected - 1)
    Call add_units_item(b_units, units.b.txt, units.b.selected - 1)
    Call add_units_item(e_units, units.e.txt, units.e.selected - 1)
    Call add_units_item(q0_units, units.q0.txt, units.q0.selected - 1)
    
    Call DisableX(frm_add_new_material)
    frm_add_new_material.Caption = frm_add_new_material.Caption & " - " & document(doc).Caption
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    If FState(doc).Conta > 1 Then
        With document(doc)
            With .lista
                .row = FState(doc).Conta - 1
                .col = 1
                l_txt.Text = .Text
                .col = 2
                area_txt.Text = .Text
                .col = 3
                k_txt.Text = .Text
                .col = 4
                b_txt.Text = .Text
                .col = 5
                te_txt.Text = .Text
                .col = 6
                td_txt.Text = .Text
                .col = 7
                n_txt.Text = .Text
                .col = 8
                alfa_txt.Text = .Text
                .col = 9
                e_txt.Text = .Text
                .col = 10
                q0_txt.Text = .Text
            End With
        End With
    End If

b_txt.Enabled = True
k_txt.Enabled = True
td_txt.Enabled = True
te_txt.Enabled = True
area_txt.Enabled = True
l_txt.Enabled = True
alfa_txt.Enabled = True
e_txt.Enabled = True
n_txt.Enabled = True
q0_txt.Enabled = True

mat_combo.Enabled = True

b_units.Locked = False

k_units.Enabled = True
td_units.Enabled = True
te_units.Enabled = True
area_units.Enabled = True
l_units.Enabled = True
q0_units.Enabled = True
e_units.Enabled = True

End Sub

' Gestao das unidades na lista de materiais
Private Sub add_units_item(name As ComboBox, ByRef arrays, ByRef default)
    Dim i As Integer
    
    For i = 1 To UBound(arrays)
        name.AddItem arrays(i)
    Next i
    name.ListIndex = default
End Sub

Private Sub b_units_Click()
    Dim i() As Double
    doc = current_form
    With document(doc)
        .lista.row = 0
        .lista.col = 4
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) * units.b.conversion(units.b.selected)
            Next j
           units.b.selected = b_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) / units.b.conversion(units.b.selected))
            Next j
           .lista.row = 0
           .lista.Text = "B (" & units.b.txt(units.b.selected) & ")"
        Else
            units.b.selected = b_units.ListIndex + 1
            .lista.row = 0
            .lista.col = 4
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "B (" & units.b.txt(units.b.selected) & ")"
        End If
    End With
End Sub

Private Sub k_units_Click()
    Dim i() As Double
    

    doc = current_form
    With document(doc)
        .lista.col = 3
        .lista.row = 0
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) * units.k.conversion(units.k.selected)
            Next j
           units.k.selected = k_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) / units.k.conversion(units.k.selected))
            Next j
            .lista.row = 0
            .lista.Text = "k (" & units.k.txt(units.k.selected) & ")"
        Else
            units.k.selected = k_units.ListIndex + 1
            .lista.col = 3
            .lista.row = 0
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "k (" & units.k.txt(units.k.selected) & ")"
        End If
    End With
End Sub

Private Sub te_units_Click()
    Dim i() As Double
    
    
    doc = current_form
    With document(doc)
        .lista.col = 5
        .lista.row = 0
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) * units.te.conversion(units.te.selected)
            Next j
           units.te.selected = te_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) / units.te.conversion(units.te.selected))
            Next j
           .lista.row = 0
           .lista.Text = "Te (" & units.te.txt(units.te.selected) & ")"
        Else
            units.te.selected = te_units.ListIndex + 1
            .lista.col = 5
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "Te (" & units.te.txt(units.te.selected) & ")"
        End If
    End With
End Sub

Private Sub td_units_Click()
    Dim i() As Double
    

    doc = current_form
    With document(doc)
        .lista.col = 6
        .lista.row = 0
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) * units.td.conversion(units.td.selected)
            Next j
           units.td.selected = td_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) / units.td.conversion(units.td.selected))
            Next j
            .lista.row = 0
            .lista.Text = "Td (" & units.td.txt(units.td.selected) & ")"
        Else
            units.td.selected = td_units.ListIndex + 1
            .lista.col = 6
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "Td (" & units.td.txt(units.td.selected) & ")"
        End If
    End With
End Sub
Private Sub l_units_Click()
    Dim i() As Double
    
    
    doc = current_form
    With document(doc)
        .lista.col = 1
        .lista.row = 0
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) * units.l.conversion(units.l.selected)
            Next j
           units.l.selected = l_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) / units.l.conversion(units.l.selected))
            Next j
           .lista.row = 0
           .lista.Text = "L (" & units.l.txt(units.l.selected) & ")"
        Else
            units.l.selected = l_units.ListIndex + 1
            .lista.col = 1
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "L (" & units.l.txt(units.l.selected) & ")"
        End If
    End With
End Sub
Private Sub area_units_Click()
    Dim i() As Double
    
    
    doc = current_form
    With document(doc)
        .lista.col = 2
        .lista.row = 0
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) * units.area.conversion(units.area.selected)
            Next j
           units.area.selected = area_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) / units.area.conversion(units.area.selected))
            Next j
            .lista.row = 0
            .lista.Text = "Area (" & units.area.txt(units.area.selected) & ")"
        Else
            units.area.selected = area_units.ListIndex + 1
            .lista.col = 2
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "Area (" & units.area.txt(units.area.selected) & ")"
        End If
    End With
End Sub

Private Sub e_units_click()
    Dim i() As Double
    
    doc = current_form
    With document(doc)
        .lista.col = 8
        .lista.row = 0
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) * units.e.conversion(units.e.selected)
            Next j
           units.e.selected = e_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) / units.e.conversion(units.e.selected))
            Next j
            .lista.row = 0
            .lista.Text = "E (" & units.e.txt(units.e.selected) & ")"
        Else
            units.e.selected = e_units.ListIndex + 1
            .lista.col = 8
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "E (" & units.e.txt(units.e.selected) & ")"
        End If
    End With

End Sub

Private Sub q0_units_click()
    Dim i() As Double
    
    doc = current_form
    With document(doc)
        .lista.col = 10
        .lista.row = 0
        .lista.CellFontBold = True
        .lista.CellAlignment = 4
        If FState(doc).Conta > 1 Then
            ReDim i(FState(doc).Conta - 1)
            For j = 1 To UBound(i)
                .lista.row = j
                i(j) = CDbl(.lista.Text) / units.q0.conversion(units.q0.selected)
            Next j
           units.q0.selected = q0_units.ListIndex + 1
            For j = 1 To UBound(i)
                .lista.row = j
                .lista.Text = CStr(i(j) * units.q0.conversion(units.q0.selected))
            Next j
            .lista.row = 0
            .lista.Text = "Q0 (" & units.q0.txt(units.q0.selected) & ")"
        Else
            units.q0.selected = q0_units.ListIndex + 1
            .lista.col = 10
            .lista.CellFontBold = True
            .lista.CellAlignment = 4
            .lista.Text = "Q0 (" & units.q0.txt(units.q0.selected) & ")"
        End If
    End With

End Sub
'fim da gestao das unidades na lista de materiais

'Gestao da alteraçao dos valores dentro das caixa de texto
Private Sub alfa_txt_Change()
  apply = False
End Sub

Private Sub area_txt_Change()
  apply = False
End Sub

Private Sub b_txt_Change()
  apply = False
End Sub

Private Sub e_txt_Change()
  apply = False
End Sub

Private Sub k_txt_Change()
  apply = False
End Sub

Private Sub l_txt_Change()
  apply = False
End Sub

Private Sub n_txt_Change()
  apply = False
End Sub

Private Sub td_txt_Change()
  apply = False
End Sub

Private Sub te_txt_Change()
  apply = False
End Sub

' fim da gestao dos valores da caixa de texto

