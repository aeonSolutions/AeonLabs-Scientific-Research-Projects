VERSION 5.00
Begin VB.Form frm_add_new_material 
   Caption         =   "material properties"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11745
   Icon            =   "frm_add_new_material.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "Apply"
      Height          =   270
      Left            =   10500
      TabIndex        =   24
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   270
      Left            =   7830
      TabIndex        =   20
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   270
      Left            =   9150
      TabIndex        =   19
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Material properties"
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   150
      TabIndex        =   0
      Top             =   330
      Width           =   11535
      Begin VB.ComboBox mat_combo 
         Height          =   315
         Left            =   4020
         TabIndex        =   36
         Top             =   390
         Width           =   2565
      End
      Begin VB.ComboBox l_units 
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
         Left            =   9510
         TabIndex        =   34
         Top             =   1650
         Width           =   1245
      End
      Begin VB.ComboBox area_units 
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
         Left            =   9510
         TabIndex        =   33
         Top             =   2130
         Width           =   1245
      End
      Begin VB.ComboBox k_units 
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
         Left            =   4920
         TabIndex        =   32
         Top             =   2070
         Width           =   1245
      End
      Begin VB.ComboBox td_units 
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
         Left            =   4920
         TabIndex        =   31
         Top             =   3000
         Width           =   1245
      End
      Begin VB.ComboBox te_units 
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
         Left            =   4920
         TabIndex        =   30
         Top             =   2490
         Width           =   1245
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
         Left            =   4920
         TabIndex        =   26
         Top             =   1590
         Width           =   1245
      End
      Begin VB.TextBox e_txt 
         Height          =   315
         Left            =   8340
         TabIndex        =   8
         Top             =   2580
         Width           =   1065
      End
      Begin VB.TextBox alfa_txt 
         Height          =   315
         Left            =   3690
         TabIndex        =   5
         Top             =   3450
         Width           =   1065
      End
      Begin VB.TextBox td_txt 
         Height          =   315
         Left            =   3690
         TabIndex        =   4
         Top             =   3000
         Width           =   1065
      End
      Begin VB.TextBox te_txt 
         Height          =   315
         Left            =   3690
         TabIndex        =   3
         Top             =   2520
         Width           =   1065
      End
      Begin VB.TextBox area_txt 
         Height          =   315
         Left            =   8340
         TabIndex        =   7
         Top             =   2130
         Width           =   1065
      End
      Begin VB.TextBox n_txt 
         Height          =   315
         Left            =   8340
         TabIndex        =   9
         Top             =   3060
         Width           =   1065
      End
      Begin VB.TextBox b_txt 
         Height          =   315
         Left            =   3675
         TabIndex        =   1
         Top             =   1590
         Width           =   1065
      End
      Begin VB.TextBox k_txt 
         Height          =   315
         Left            =   3690
         TabIndex        =   2
         Top             =   2070
         Width           =   1065
      End
      Begin VB.TextBox l_txt 
         Height          =   315
         Left            =   8340
         TabIndex        =   6
         Top             =   1650
         Width           =   1065
      End
      Begin VB.Label tipo_txt 
         Height          =   375
         Left            =   1410
         TabIndex        =   35
         Top             =   360
         Width           =   2325
      End
      Begin VB.Label Label14 
         Caption         =   "GPa"
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
         Left            =   9570
         TabIndex        =   29
         Top             =   2640
         Width           =   525
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
         Left            =   9600
         TabIndex        =   28
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
         Left            =   4980
         TabIndex        =   27
         Top             =   1230
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   1950
         Left            =   330
         Picture         =   "frm_add_new_material.frx":08CA
         Top             =   1650
         Width           =   2250
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
         Left            =   5490
         TabIndex        =   25
         Top             =   3420
         Width           =   315
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   6810
         X2              =   6810
         Y1              =   1230
         Y2              =   3930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   6840
         X2              =   6840
         Y1              =   1230
         Y2              =   3900
      End
      Begin VB.Label Label11 
         Caption         =   "E :"
         Height          =   255
         Left            =   7920
         TabIndex        =   23
         Top             =   2610
         Width           =   315
      End
      Begin VB.Label Label10 
         Caption         =   "/ºC x10"
         Height          =   225
         Left            =   4950
         TabIndex        =   22
         Top             =   3510
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
         Left            =   3270
         TabIndex        =   21
         Top             =   3450
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "Td :"
         Height          =   225
         Left            =   3210
         TabIndex        =   18
         Top             =   3030
         Width           =   405
      End
      Begin VB.Label Label17 
         Caption         =   "Te :"
         Height          =   225
         Left            =   3210
         TabIndex        =   17
         Top             =   2580
         Width           =   405
      End
      Begin VB.Label Label9 
         Caption         =   "Area :"
         Height          =   225
         Left            =   7710
         TabIndex        =   16
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "Split :"
         Height          =   255
         Left            =   7770
         TabIndex        =   15
         Top             =   3090
         Width           =   465
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
         Left            =   3300
         TabIndex        =   14
         Top             =   1590
         Width           =   270
      End
      Begin VB.Label Label2 
         Caption         =   "k :"
         Height          =   270
         Left            =   3300
         TabIndex        =   13
         Top             =   2070
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "L :"
         Height          =   225
         Left            =   7920
         TabIndex        =   12
         Top             =   1650
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Phisical Properties:"
         Height          =   225
         Left            =   7380
         TabIndex        =   11
         Top             =   1290
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Thermal properties:"
         Height          =   315
         Left            =   2430
         TabIndex        =   10
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
Dim lastindex As Integer
Dim apply As Boolean

Public doc As Integer

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
    If tipo = "new" Then
       lastindex = FState(document(doc).Tag).Conta
       FState(doc).Conta = FState(doc).Conta + 1
       mat_combo.Clear
       For j = 1 To FState(doc).Conta
         mat_combo.AddItem "Material nº" & CStr(j)
       Next j
       mat_combo.ListIndex = FState(doc).Conta - 1
    End If

    FState(document(doc).Tag).saved = False
    With document(doc)
    If .lista.Rows <= lastindex Then
        .lista.Rows = lastindex + 1
    End If
    .lista.col = 1
    .lista.row = lastindex
    .lista.col = 0
    .lista.CellAlignment = 4
    .lista.Text = Str(lastindex)
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
    .lista.col = 2
    .lista.CellAlignment = 4
    .lista.Text = area_txt.Text
    .lista.col = 4
    .lista.CellAlignment = 4
    .lista.Text = b_txt.Text
    For j = 1 To lastindex
        .lista.row = j
        .lista.col = 5
        .lista.CellAlignment = 4
        .lista.Text = te_txt.Text
        .lista.col = 6
        .lista.CellAlignment = 4
        .lista.Text = td_txt.Text
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
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Not apply Then
        Call CmdApply_Click
    End If
    Unload Me

End Sub
Private Sub mat_combo_Click()
If mat_combo.ListIndex <> lastindex Then
    On Error Resume Next
    i = current_form
    lastindex = mat_combo.ListIndex
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
        End With
    End With
End If
End Sub

Private Sub form_load()
    Dim tmp As String
    Dim vector() As String
    
    doc = current_form
    mat_combo.Clear
    If tipo = "new" Then
       For j = 1 To FState(doc).Conta
         mat_combo.AddItem "Material nº" & CStr(j)
       Next j
       mat_combo.ListIndex = FState(doc).Conta - 1
       mat_combo.Enabled = False
       tipo_txt = "Inserting material : "
    ElseIf tipo = "edit" Then
       For j = 1 To FState(doc).Conta - 1
         mat_combo.AddItem "Material nº" & CStr(j)
       Next j
       mat_combo.ListIndex = 0
       mat_combo.Enabled = True
       tipo_txt = "Material to edit:"
    End If
    ReDim vector(2)
    vector(1) = "W/m2"
    vector(2) = "W/mm2.ºC"
    Call add_units_item(b_units, vector, 1)
    vector(1) = "W/m.ºC"
    vector(2) = "W/mm.ºC"
    Call add_units_item(k_units, vector, 1)
    vector(1) = "ºC"
    vector(2) = "ºK"
    Call add_units_item(te_units, vector, 0)
    vector(1) = "ºC"
    vector(2) = "ºK"
    Call add_units_item(td_units, vector, 1)
    ReDim vector(5)
    vector(1) = "um"
    vector(2) = "mm"
    vector(3) = "cm"
    vector(4) = "dm"
    vector(5) = "m"
    Call add_units_item(l_units, vector, 1)
    vector(1) = "um2"
    vector(2) = "mm2"
    vector(3) = "cm2"
    vector(4) = "dm2"
    vector(5) = "m2"
    Call add_units_item(area_units, vector, 1)
    
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
            End With
        End With
    End If
b_units.Enabled = False
k_units.Enabled = False
l_units.Enabled = False
area_units.Enabled = False
td_units.Enabled = False
te_units.Enabled = False

End Sub

' Gestao das unidades na lista de materiais
Private Sub add_units_item(name As ComboBox, ByRef arrays, ByRef default)
    Dim i As Integer
    
    For i = 1 To UBound(arrays)
        name.AddItem arrays(i)
    Next i
    name.ListIndex = default
End Sub

Private Sub q0_units_Click()
    doc = current_form
    If q0_units.ListIndex = 0 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 4
        document(doc).lista.Text = "q0 (W/m3)"
    Else
        document(doc).lista.row = 0
        document(doc).lista.col = 4
        document(doc).lista.Text = "qo (W/mm3)"
    End If
End Sub

Private Sub k_units_Click()
    doc = current_form
    If k_units.ListIndex = 0 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 3
        document(doc).lista.Text = "k (W/m.ºC)"
    Else
        document(doc).lista.row = 0
        document(doc).lista.col = 3
        document(doc).lista.Text = "k (W/mm.ºC)"
    End If
End Sub

Private Sub te_units_Click()
    doc = current_form
    If te_units.ListIndex = 0 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 5
        document(doc).lista.Text = "Te (ºC)"
    Else
        document(doc).lista.row = 0
        document(doc).lista.col = 5
        document(doc).lista.Text = "Te (ºK)"
    End If
End Sub
Private Sub l_units_Click()
    doc = current_form
    If l_units.ListIndex = 0 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 1
        document(doc).lista.Text = "l (um)"
    ElseIf l_units.ListIndex = 1 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 1
        document(doc).lista.Text = "l (mm)"
    ElseIf l_units.ListIndex = 2 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 1
        document(doc).lista.Text = "l (cm)"
    ElseIf l_units.ListIndex = 3 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 1
        document(doc).lista.Text = "l (dm)"
    ElseIf l_units.ListIndex = 4 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 1
        document(doc).lista.Text = "l (m)"
    End If
End Sub
Private Sub area_units_Click()
    doc = current_form
    If area_units.ListIndex = 0 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 2
        document(doc).lista.Text = "area (um2)"
    ElseIf l_units.ListIndex = 1 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 2
        document(doc).lista.Text = "area (mm2)"
    ElseIf l_units.ListIndex = 2 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 2
        document(doc).lista.Text = "area (cm2)"
    ElseIf l_units.ListIndex = 3 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 2
        document(doc).lista.Text = "area (dm2)"
    ElseIf l_units.ListIndex = 4 Then
        document(doc).lista.row = 0
        document(doc).lista.col = 2
        document(doc).lista.Text = "area (m2)"
    End If
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

