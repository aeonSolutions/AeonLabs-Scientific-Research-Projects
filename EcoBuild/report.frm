VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_report_analysis 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Analysys Report"
   ClientHeight    =   8130
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   14430
   Icon            =   "report.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleMode       =   0  'User
   ScaleWidth      =   14430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dialogs 
      Left            =   270
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox 
      Height          =   8025
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   14155
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"report.frx":2052
   End
   Begin VB.Menu viewother 
      Caption         =   "&View graphs"
      Begin VB.Menu mnu_global 
         Caption         =   "Global Analysis"
      End
      Begin VB.Menu mnu_energy 
         Caption         =   "Energy Consuption"
      End
      Begin VB.Menu mnu_water 
         Caption         =   "Water Consuption"
      End
      Begin VB.Menu mnu_nox 
         Caption         =   "NOx emissions"
      End
      Begin VB.Menu mnu_co2 
         Caption         =   "CO2 emissions"
      End
      Begin VB.Menu mnu_so2 
         Caption         =   "Structure Costs"
      End
   End
   Begin VB.Menu printgraph 
      Caption         =   "&Print"
   End
   Begin VB.Menu Save 
      Caption         =   "&Save Report"
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frm_report_analysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim doc As Integer
Dim tmp As String
Dim tmp2 As String
Dim i As Integer

doc = current_form
Call DisableX(frm_report(doc))
RichTextBox.LoadFile App.path & "\report.rtf"
tmp = RichTextBox.TextRTF
tmp = Replace(tmp, "[CODED]", FState(doc).name, 1, 1, vbTextCompare)
document(doc).lista.Row = 1
document(doc).lista.Col = 18
tmp = Replace(tmp, "[CODED]", document(doc).lista.Text, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).cement, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).aggregates, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).concrete_cost, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).metal_cost, 1, 1, vbTextCompare)
If doc_props(doc).impact_concrete.costs > doc_props(doc).impact_metal.costs Then
    tmp2 = "less"
ElseIf doc_props(doc).impact_concrete.costs < doc_props(doc).impact_metal.costs Then
    tmp2 = "more"
Else
    tmp2 = "the same"
End If
tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).cement_qty, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).armour_qty, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).aggregates_qty, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).total_weight, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).volume_concrete, 1, 1, vbTextCompare)
tmp = Replace(tmp, "[CODED]", doc_props(doc).volume_wood, 1, 1, vbTextCompare)
With doc_props(doc).impact_metal
    tmp = Replace(tmp, "[CODED]", .energy, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .water, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .co2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .so2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .nox, 1, 1, vbTextCompare)
End With
With doc_props(doc).impact_transport
    tmp = Replace(tmp, "[CODED]", .co2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .so2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .nox, 1, 1, vbTextCompare)
End With
With doc_props(doc).impact_concrete
    tmp = Replace(tmp, "[CODED]", .energy, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .water, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .co2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .so2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .nox, 1, 1, vbTextCompare)
End With
With doc_props(doc).impact_total
    tmp = Replace(tmp, "[CODED]", .energy, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .water, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .co2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .so2, 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", .nox, 1, 1, vbTextCompare)
End With
tmp2 = ""
For i = 1 To FState(doc).count
    tmp2 = tmp2 & "\intbl "
    With document(doc).lista
        .Row = i
        .Col = 0
        tmp2 = tmp2 & .Text & "\cell " ' type of pillar/beam
        .Col = 1
        tmp2 = tmp2 & .Text & "\cell " ' quantity
        .Col = 2
        tmp2 = tmp2 & .Text & "\cell " ' height
        .Col = 3
        tmp2 = tmp2 & .Text & "\cell " ' width
        .Col = 4
        tmp2 = tmp2 & .Text & "\cell " ' weight
        .Col = 5
        tmp2 = tmp2 & .Text & "\cell " ' lenght
        .Col = 6
        tmp2 = tmp2 & .Text & "\cell " ' f5
        .Col = 7
        tmp2 = tmp2 & .Text & "\cell " ' f6
        .Col = 8
        tmp2 = tmp2 & .Text & "\cell " ' f8
        .Col = 9
        tmp2 = tmp2 & .Text & "\cell " ' f10
        .Col = 10
        tmp2 = tmp2 & .Text & "\cell " ' f12
        .Col = 11
        tmp2 = tmp2 & .Text & "\cell " ' f16
        .Col = 12
        tmp2 = tmp2 & .Text & "\cell " ' f20
        .Col = 13
        tmp2 = tmp2 & .Text & "\cell " ' f25
        .Col = 14
        tmp2 = tmp2 & .Text & "\cell" ' f32
    End With
    tmp2 = tmp2 & "\row " & vbCrLf
Next i
tmp = Replace(tmp, "\intbl\cell\cell\cell\cell\cell\cell\cell\cell\cell\cell\cell\cell\cell\cell\cell\row", tmp2, 1, 1, vbTextCompare)

RichTextBox.TextRTF = tmp
End Sub

Private Sub exit_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).Hide
Unload Me
End Sub


Private Sub mnu_co2_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).Hide
Unload Me
frm_co2(doc).Show 1


End Sub

Private Sub mnu_energy_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).Hide
Unload Me
frm_energy(doc).Show 1

End Sub

Private Sub mnu_global_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).Hide
Unload Me
frm_global(doc).Show 1

End Sub

Private Sub mnu_nox_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).Hide
Unload Me
frm_nox(doc).Show 1

End Sub

Private Sub mnu_so2_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).Hide
Unload Me
frm_structure(doc).Show 1

End Sub

Private Sub mnu_water_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).Hide
Unload Me
frm_water(doc).Show 1

End Sub

Private Sub printgraph_Click()
Dim doc As Integer
doc = current_form

frm_report(doc).RichTextBox.SelPrint (Printer.hDC)
End Sub

Private Sub Save_Click()
' Set CancelError is True
  dialogs.CancelError = True
 On Error Resume Next
  ' Set flags
  dialogs.Flags = cdlOFNHideReadOnly And cdlOFNAllowMultiselect
  ' Set filters
  dialogs.Filter = "Rich Text Files" & "(*.rtf)|*.rtf"
  ' Specify default filter
  dialogs.FilterIndex = 2
  ' Display the open dialog box
   dialogs.ShowSave
  If Err.Number = 32755 Then ' cancel was selected
    'MsgBox "error num:" & Str(Err.Number) & " Desc:" & Err.Description, vbOKCancel, "Info"
    Exit Sub
  End If
  ' get the name file and the path
   RichTextBox1.SaveFile dialogs.filename, rtfRTF
End Sub
