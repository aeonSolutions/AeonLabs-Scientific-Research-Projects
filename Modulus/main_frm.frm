VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form main_frm 
   Caption         =   "Modulus"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   1050
   ClientWidth     =   10995
   Icon            =   "main_frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Material Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3585
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10815
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
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
         Left            =   8760
         TabIndex        =   8
         Top             =   1020
         Width           =   1125
      End
      Begin VB.CommandButton insert 
         Caption         =   "Insert"
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
         Left            =   8760
         TabIndex        =   7
         Top             =   540
         Width           =   1125
      End
      Begin VB.TextBox e_txt 
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
         Height          =   255
         Left            =   6960
         TabIndex        =   6
         Text            =   "?"
         Top             =   1110
         Width           =   1065
      End
      Begin VB.TextBox u_txt 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   2700
         Width           =   1065
      End
      Begin VB.TextBox a_txt 
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
         Height          =   255
         Left            =   2910
         TabIndex        =   3
         Top             =   2700
         Width           =   1065
      End
      Begin VB.TextBox b_txt 
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
         Height          =   255
         Left            =   6960
         TabIndex        =   4
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox h_txt 
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
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   720
         Width           =   1065
      End
      Begin VB.TextBox f_txt 
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
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   2700
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   1215
         Left            =   5460
         TabIndex        =   13
         Top             =   1950
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   15
         Cols            =   4
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin VB.Label Label12 
         Caption         =   "a (mm)"
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
         Left            =   3120
         TabIndex        =   25
         Top             =   2490
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "u (mm)"
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
         Left            =   1800
         TabIndex        =   24
         Top             =   2490
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "F (N)"
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
         Left            =   570
         TabIndex        =   23
         Top             =   2460
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   4740
         Picture         =   "main_frm.frx":08CA
         Top             =   330
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "note: substrate material properties must be the first entry (S1)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5340
         TabIndex        =   20
         Top             =   3210
         Width           =   4935
      End
      Begin VB.Label Label5 
         Caption         =   "b:"
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
         Left            =   6630
         TabIndex        =   19
         Top             =   390
         Width           =   285
      End
      Begin VB.Label Label4 
         Caption         =   "h:"
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
         Left            =   6630
         TabIndex        =   18
         Top             =   750
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "E:"
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
         Left            =   6630
         TabIndex        =   17
         Top             =   1140
         Width           =   285
      End
      Begin VB.Label Label2 
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
         Left            =   8130
         TabIndex        =   16
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "mm2"
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
         Left            =   8130
         TabIndex        =   15
         Top             =   540
         Width           =   525
      End
      Begin VB.Image Image1 
         Height          =   1950
         Left            =   360
         Picture         =   "main_frm.frx":0DB9
         Top             =   360
         Width           =   3825
      End
      Begin VB.Line Line1 
         X1              =   4320
         X2              =   4320
         Y1              =   450
         Y2              =   2970
      End
   End
   Begin VB.Frame frm_output 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   3870
      Width           =   10815
      Begin VB.TextBox iexp_txt 
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
         Height          =   255
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   1065
      End
      Begin MSFlexGridLib.MSFlexGrid results 
         Height          =   1215
         Left            =   5160
         TabIndex        =   12
         Top             =   420
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   15
         Cols            =   4
         ScrollBars      =   2
         Appearance      =   0
      End
      Begin VB.TextBox i_teo_txt 
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
         Height          =   255
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1200
         Width           =   1065
      End
      Begin VB.TextBox z_txt 
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
         Height          =   255
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label13 
         Caption         =   "experimental inertia"
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
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "note: results aproximated to 0,001"
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
         Left            =   3000
         TabIndex        =   26
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "theoretical inertia"
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
         TabIndex        =   22
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label Label8 
         Caption         =   "neutral axis"
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
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   1155
      End
   End
   Begin VB.Label Label6 
      Caption         =   "www.MiguelSilva.web.pt"
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
      Height          =   285
      Left            =   8220
      TabIndex        =   14
      Top             =   5940
      Width           =   2445
   End
   Begin VB.Menu menu_calc 
      Caption         =   "&Calculate"
   End
   Begin VB.Menu menu_exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "main_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lista_count As Integer
Public incognit As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub clear_Click()
Dim tmp As VbMsgBoxResult

  tmp = MsgBox("Clear all section elements ?", vbYesNo + vbCritical, " Modulus ")
  If tmp = vbCancel Then
    Exit Sub
  End If
  If tmp = vbYes Then
    lista.clear
    lista.Row = 0
    lista.Col = 0
    lista.CellAlignment = 4
    lista.CellFontBold = True
    lista.Text = "Si"
    lista.Col = 1
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "b (m) "
    lista.Col = 2
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "h (m) "
    lista.Col = 3
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "E (GPa) "
    lista_count = 0
  End If
  incognit = False
End Sub
Private Sub verify_inertia()
If IsNumeric(f_txt) And IsNumeric(u_txt) And IsNumeric(a_txt) And lista_count > 0 Then
  lista.Row = 1
  lista.Col = 3
  iexp_txt.Text = CStr(Round(4 * CDbl(f_txt.Text) * CDbl(a_txt.Text) ^ 3 / (3 * CDbl(u_txt.Text) * CDbl(lista.Text) * 1000), 8))
End If
End Sub

Private Sub Form_Load()
    With lista
        .ColWidth(0) = TextWidth("###") * 2
        .ColWidth(1) = TextWidth("######") * 2
        .ColWidth(2) = TextWidth("######") * 2
        .ColWidth(3) = TextWidth("######") * 2
        .Row = 0
        .Col = 0
        .CellAlignment = 4
        .CellFontBold = True
        .Text = "Si"
        .Col = 1
        .CellFontBold = True
        .CellAlignment = 4
        .Text = "b (mm) "
        .Col = 2
        .CellFontBold = True
        .CellAlignment = 4
        .Text = "h (mm) "
        .Col = 3
        .CellFontBold = True
        .CellAlignment = 4
        .Text = "E (GPa) "
    End With
    With results
        .ColWidth(0) = TextWidth("###") * 2
        .ColWidth(1) = TextWidth("######") * 2
        .ColWidth(2) = TextWidth("######") * 2
        .ColWidth(3) = TextWidth("######") * 2
        .Row = 0
        .Col = 0
        .CellAlignment = 4
        .CellFontBold = True
        .Text = "Si"
        .Col = 1
        .CellFontBold = True
        .CellAlignment = 4
        .Text = "m"
        .Col = 2
        .CellFontBold = True
        .CellAlignment = 4
        .Text = "b' (mm) "
        .Col = 3
        .CellFontBold = True
        .CellAlignment = 4
        .Text = "E (GPa)"
    End With
    lista_count = 0
    incognit = False



End Sub

Private Sub Insert_Click()
   If Not IsNumeric(b_txt) Then
        b_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(h_txt) Then
        h_txt.SetFocus
        Exit Sub
    End If
   If Not IsNumeric(e_txt) And e_txt.Text <> "?" Then
        e_txt.SetFocus
        Exit Sub
    End If
    If lista_count = 0 And e_txt.Text = "?" Then
         MsgBox "Substrate material properties must be the first entry[S1] and E cannot be an incognit[?]", vbOK + vbCritical, " Temperus "
        Exit Sub
    End If
    If incognit = True And e_txt.Text = "?" Then
        MsgBox "There's already one Incognit in the Elastic modulus", vbOK + vbCritical, " Temperus "
        Exit Sub
    End If
    If e_txt.Text = "?" And incognit = False Then
        incognit = True
    End If
lista_count = lista_count + 1
If lista_count >= lista.Rows Then
    lista.Rows = lista_count + 1
End If
With lista
    .Row = lista_count
    .Col = 0
    .CellAlignment = 4
    .Text = "S" & Str(lista_count)
    .Col = 1
    .CellAlignment = 4
    .Text = b_txt.Text
    .Col = 2
    .CellAlignment = 4
    .Text = h_txt.Text
    .Col = 3
    .CellAlignment = 4
    .Text = e_txt.Text
End With
Call verify_inertia
e_txt.Text = "?"
End Sub


Private Sub menu_Calc_Click()
Dim i As Integer
Dim j As Integer
Dim tj As Double
Dim b() As Double
Dim b_in() As Double
Dim h() As Double
Dim m() As Double
Dim e() As Double
Dim z As Double
Dim esub As Double
Dim sum_par As Double
Dim sum_par_2 As Double
Dim i_exp As Double
Dim i_teo As Double
Dim missing As Integer
Dim k As Integer


If Not IsNumeric(f_txt) Then
    f_txt.SetFocus
    Exit Sub
End If
If Not IsNumeric(u_txt) Then
    u_txt.SetFocus
    Exit Sub
End If
If Not IsNumeric(a_txt) Then
    a_txt.SetFocus
    Exit Sub
End If
If lista_count = 0 Then
    b_txt.SetFocus
    Exit Sub
End If
If incognit = False Then
    MsgBox "It should be at least one unknown in the elastic modulus", vbOK + vbCritical, " Temperus "
    Exit Sub
End If
If Not IsNumeric(iexp_txt) Then
    Call verify_inertia
End If
If lista_count >= results.Rows Then
    results.Rows = lista_count + 1
End If
i_exp = CDbl(iexp_txt.Text)
ReDim m(lista_count)
ReDim b(lista_count)
ReDim e(lista_count)
ReDim h(lista_count)
ReDim b_in(lista_count)

lista.Col = 3
lista.Row = 1
'modulo elasticidade do substrato
i_teo = 9.99999E+20
j = 0
For i = 1 To lista_count
    lista.Col = 3
    lista.Row = i
    If lista.Text = "?" Then
        j = i
    Else
        e(i) = CDbl(lista.Text)
        If e(i) < i_teo Then
            i_teo = e(i)
        End If
    End If
Next i
e(j) = i_teo
missing = j
i_teo = -9999999999#
esub = e(1)
For i = 1 To lista_count
        lista.Row = i
        lista.Col = 2
        h(i) = CDbl(lista.Text)
        lista.Col = 1
        b_in(i) = CDbl(lista.Text)
Next i

Do
    'determinar o factor de homogenizaçao e a respectiva largura b
    For i = 1 To lista_count
        m(i) = e(i) / esub
        b(i) = m(i) * b_in(i)
    Next i
    'determinar o eixo neutro
    sum_par = 0
    For i = 1 To lista_count
        tj = 0
        For j = 1 To i - 1
            tj = tj + h(j)
        Next j
        sum_par = sum_par + b(i) * h(i) * (2 * tj + h(i))
    Next i
    sum_par_2 = 0
    For i = 1 To lista_count
        sum_par_2 = sum_par_2 + b(i) * h(i)
    Next i
    z = sum_par / (2 * sum_par_2)
    'determinar a inercia
    i_teo = 0
    For i = 1 To lista_count
        sum_par = 0
        For j = 1 To i
            sum_par = sum_par + h(j)
        Next j
        sum_par_2 = 0
        For j = 1 To i - 1
            sum_par_2 = sum_par_2 + h(j)
        Next j
        i_teo = i_teo + b(i) * ((sum_par - z) ^ 3 + (z - sum_par_2) ^ 3)
    Next i
    i_teo = i_teo / 3
    If i_exp - 0.001 < i_teo And i_teo < i_exp + 0.001 Then
        Exit Do
    End If
    If i_exp - 0.001 > i_teo Then
        e(missing) = e(missing) + e(1) * 0.001
    End If
    If i_exp + 0.001 < i_teo Then
        e(missing) = e(missing) - e(1) * 0.001
    End If
Loop
z_txt.Text = Str(Round(z, 8))
i_teo_txt.Text = Str(Round(i_teo, 8))
For i = 1 To lista_count
        With results
            .Row = i
            .Col = 0
            .CellAlignment = 4
            .Text = Str(i)
            .Col = 1
            .CellAlignment = 4
            .Text = CStr(Round(m(i), 5))
            .Col = 2
            .CellAlignment = 4
            .Text = CStr(Round(b(i), 5))
            .Col = 3
            .CellAlignment = 4
            .Text = CStr(Round(e(i), 5))
        End With
Next i

End Sub

Private Sub menu_exit_Click()
Unload Me
End Sub

Private Sub u_txt_Change()
 Call verify_inertia
End Sub

Private Sub a_txt_Change()
 Call verify_inertia
End Sub

Private Sub f_txt_Change()
 Call verify_inertia
End Sub

Private Sub Label6_Click()
    ShellExecute Me.hwnd, vbNullString, "http://www.miguelsilva.web.pt", vbNullString, "C:\", 1

End Sub

