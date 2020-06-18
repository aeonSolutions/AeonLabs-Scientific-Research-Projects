VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_graph_exp_data 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8130
   ClientLeft      =   -480
   ClientTop       =   2745
   ClientWidth     =   14820
   Icon            =   "frm_graph_exp_data.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleMode       =   0  'User
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Begin MSChart20Lib.MSChart chart_exp_data 
      Height          =   7575
      Left            =   90
      OleObjectBlob   =   "frm_graph_exp_data.frx":08CA
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   14625
   End
   Begin VB.Label info_txt 
      Caption         =   "Plotting at [CODED] %"
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   2595
   End
   Begin VB.Menu mnu_selection 
      Caption         =   "&Selection"
      Begin VB.Menu mnu_extension 
         Caption         =   "&Extension"
      End
   End
   Begin VB.Menu printgraph 
      Caption         =   "&Print"
   End
   Begin VB.Menu CopyGraph 
      Caption         =   "&Copy"
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frm_graph_exp_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim HeightDiff As Integer
Dim WidthDiff As Integer

Private Const FORMHEIGHT = 8865
Private Const FORMWIDTH = 14940
Public chht As Long
Public chwd As Long
Dim chht1 As Long
Dim chwd1 As Long

Sub form_activate()
On Error Resume Next
If Me.Width < FORMWIDTH Then Me.Width = FORMWIDTH
If Err <> 0 Then
    Exit Sub
End If

Call load_data
End Sub
Sub load_data()
Dim doc As Integer
Dim delta As Double
Dim vector() As Double
Dim i As Integer
Dim label_x As String
Dim label_y As String
Dim num_columns As Integer
Dim num_rows As Integer
Dim points As Integer
Dim tmp As Double
Dim k As Integer
Dim chart_type As Variant
Dim clabel() As String
Dim maxx As Double

doc = current_form
Call DisableX(frm_exp_data(doc))
Call Centerform(frm_exp_data(doc))

points = Round(doc_props(doc).exp_data.emax / doc_props(doc).exp_data.delta_e, 0)
ReDim vector(points, 1 To 2)
Me.Caption = graph_type
delta = doc_props(doc).exp_data.delta_e
num_columns = 2
num_rows = 1
chart_type = VtChChartType2dXY

With doc_props(doc)
    If graph_type = "Composit Stress Curve" Then
        mnu_selection.Enabled = False
        info_txt.Caption = ""
        ReDim clabel(num_columns)
        clabel(1) = "S(e[%])"
        label_x = "Mpa"
        label_y = "e (%)"
        
        For i = 0 To points
            vector(i, 1) = i * .exp_data.delta_e * 100
            With .stress_c_curve
                vector(i, 2) = .x5 * (i * delta) ^ 5 + .x4 * (i * delta) ^ 4 + .x3 * (i * delta) ^ 3 + .x2 * (i * delta) ^ 2 + .x * (i * delta) + .c
                vector(i, 2) = vector(i, 2) / 1000000 ' converted from Pa to MPa
            End With
        Next i
    ElseIf graph_type = "Polynomial Load Curve" Then
        mnu_selection.Enabled = False
        info_txt.Caption = ""
        label_x = "KN"
        label_y = "e (%)"
        ReDim clabel(num_columns)
        clabel(1) = "F(e[%])"
        
        For i = 0 To points
            vector(i, 1) = i * .exp_data.delta_e * 100
            With .exp_data
                vector(i, 2) = .x5 * (i * delta) ^ 5 + .x4 * (i * delta) ^ 4 + .x3 * (i * delta) ^ 3 + .x2 * (i * delta) ^ 2 + .x * (i * delta) + .c
                tmp = vector(i, 2)
            End With
        Next i
    ElseIf graph_type = "Composit Elastic Modulus Curve" Then
        mnu_selection.Enabled = False
        info_txt.Caption = ""
        label_x = "GPa"
        label_y = "e (%)"
        ReDim clabel(num_columns)
        clabel(1) = "E(e[%])"
        
        For i = 0 To points
            vector(i, 1) = i * .exp_data.delta_e * 100
            With .modulus_c_curve
                vector(i, 2) = .x4 * (i * delta) ^ 4 + .x3 * (i * delta) ^ 3 + .x2 * (i * delta) ^ 2 + .x * (i * delta) + .c
                vector(i, 2) = vector(i, 2) / 1000000000#
            End With
        Next i
    ElseIf graph_type = "Final Stresses" Then
        info_txt.Caption = "Plotting at " & Str(Round(.exp_data.emax / .exp_data.delta_e, 0) * .exp_data.delta_e * 100) & " %"
        mnu_selection.Enabled = False
        label_x = "MPa"
        label_y = "element"
        ReDim clabel(num_columns)
        clabel(1) = "S(element)"
        
        ReDim vector(1 To .statistic.elements, 1 To 2)
        For i = 1 To .statistic.elements
          If .elements(i).cracked Then
              vector(i, 2) = 0
          Else
              vector(i, 2) = CDbl(.elements(i).sigma / 1000000#)
          End If
          vector(i, 1) = i
        Next i
    ElseIf graph_type = "Loading Stresses" Then
        info_txt.Caption = "Plotting at " & Str(.results.live_data_pos * .exp_data.delta_e * 100) & " %"
        label_x = "MPa"
        label_y = "element"
        ReDim clabel(num_columns)
        clabel(1) = "S(element)"
        ReDim vector(1 To .statistic.elements, 1 To 2)
        With .results.live_data(.results.live_data_pos)
            For i = 1 To doc_props(doc).statistic.elements
              vector(i, 2) = CDbl(.sigma(i) / 1000000#)
              vector(i, 1) = i
            Next i
        End With
    ElseIf graph_type = "Lmed versus Stress (Homogen.)" Then
        mnu_selection.Enabled = False
        info_txt.Caption = ""
        label_x = "MPa"
        label_y = "L= " & CStr(.phisical.lenght * 1000) & " mm"
        num_columns = UBound(.results.graphics)
        num_rows = 100
        ReDim clabel(num_columns)
        For i = 1 To UBound(.results.graphics)
            clabel(i) = "e=" & CStr(.results.graphics(i).strain) & "%"
        Next i
        chart_type = VtChChartType2dLine
        ReDim vector(1 To 100, 1 To UBound(.results.graphics))
        For i = 1 To UBound(.results.graphics)
            For k = 1 To 100
                vector(k, i) = .results.graphics(i).hl(k) / 1000 ' converted to mm
                vector(k, i) = .results.graphics(i).hsf(k) / 1000000 ' converted to MPa
            Next k
        Next i
    ElseIf graph_type = "Lmed versus Stress" Then
        mnu_selection.Enabled = False
        info_txt.Caption = ""
        label_x = "MPa"
        label_y = "Lmed (mm)"
        num_columns = UBound(.results.graphics)
        num_rows = .results.graphics(1).max_elements
        ReDim clabel(num_columns)
        For i = 1 To UBound(.results.graphics)
            clabel(i) = "Lmed=" & CStr(Round(.results.graphics(i).l_med * 1000, 3)) & "mm @ " & CStr(.results.graphics(i).strain) & " %"
            
        Next i
        chart_type = VtChChartType2dLine
        ReDim vector(1 To .results.graphics(1).max_elements, 1 To UBound(.results.graphics))
        For i = 1 To UBound(.results.graphics)
            For k = 1 To .results.graphics(1).max_elements
                vector(k, i) = .results.graphics(i).l(k) / 1000 'converted to mm
                vector(k, i) = .results.graphics(i).sf(k) / 1000000 ' converted to MPa
            Next k
        Next i
    ElseIf graph_type = "Stress(Lmed) versus Strain" Then
        mnu_selection.Enabled = False
        info_txt.Caption = ""
        label_x = "MPa"
        label_y = "e (%)"
        ReDim clabel(num_columns)
        clabel(1) = "Lmed(e[%])"
        ReDim vector(1 To UBound(.results.graphics), 1 To 2)
        For i = 1 To UBound(.results.graphics)
          vector(i, 2) = .results.graphics(i).strain_l_med / 1000000
          vector(i, 1) = .results.graphics(i).strain
        Next i
        With frm_exp_data(doc).chart_exp_data.Plot.SeriesCollection(1).DataPoints(-1).Marker
            .Visible = True
            .Style = VtMarkerStyleStar
            .Pen.Style = VtPenStyleSolid
        End With
    ElseIf graph_type = "Lmed and Crk density versus Strain" Then
        mnu_selection.Enabled = False
        info_txt.Caption = ""
        label_x = "Crack Density (crks /m)"
        label_y = "e(%)"
        num_columns = 3
        num_rows = UBound(.results.graphics)
        ReDim clabel(num_columns)
'        clabel(1) = "e(%)"
        clabel(1) = "Crack Density (crk /m)"
        clabel(2) = "Lmed (mm)"
        chart_type = VtChChartType2dLine
        ReDim vector(1 To UBound(.results.graphics), 1 To 3)
        maxx = -999999999
        For i = 1 To UBound(.results.graphics)
            vector(i, 1) = .results.graphics(i).strain
            vector(i, 2) = .results.graphics(i).crk_density
            vector(i, 3) = Round(.results.graphics(i).l_med * 1000, 3)
            If vector(i, 3) > maxx Then
                maxx = vector(i, 3)
            End If
        Next i
        With frm_exp_data(doc).chart_exp_data
            .chartType = chart_type
            .ChartData = vector()
            .RowCount = num_rows
            For i = 1 To num_rows
                .Row = i
                .RowLabelIndex = 1
                .RowLabel = CStr(vector(i, 1)) & "%"
            Next i
            .ColumnCount = num_columns
            .ColumnLabelCount = num_columns
            .Column = 2
            .ColumnLabelIndex = 1
            .ColumnLabel = clabel(1)
            .Column = 3
            .ColumnLabelIndex = 1
            .ColumnLabel = clabel(2)
            .ShowLegend = False
            .Legend.Location.LocationType = VtChLocationTypeTop
            With .Plot
                .SeriesCollection(1).StatLine.flag = 0
                
                .SeriesCollection(1).SecondaryAxis = False
                .SeriesCollection(2).SecondaryAxis = False
                .SeriesCollection(3).SecondaryAxis = True
                .SeriesCollection(2).Pen.VtColor.Set 0, 0, 255 ' crack density -red,green,blue
                .SeriesCollection(3).Pen.VtColor.Set 0, 180, 0 ' Lmed
                
                With .Axis(VtChAxisIdY2)
                    .AxisTitle = clabel(2)
                    .AxisTitle.VtFont.VtColor.Set 0, 180, 0
                    .AxisScale.Hide = False
                    .CategoryScale.Auto = True
                    .ValueScale.Maximum = maxx
                    .ValueScale.Minimum = 0
                End With
                With .Axis(VtChAxisIdY)
                    .AxisTitle.VtFont.VtColor.Set 0, 0, 255
                    .AxisTitle = label_x
                    .AxisScale.Hide = False
                    .CategoryScale.Auto = True
                End With
                .UniformAxis = False
            End With
            .Visible = True
            .Refresh
        End With
        Exit Sub
ElseIf graph_type = "Lmed versus Stress(Lmed)" Then
        info_txt.Caption = ""
        mnu_selection.Enabled = False
        label_x = "S(Lmed) MPa"
        
        label_y = "Lmed (um) (logarithmic)"
        ReDim clabel(num_columns)
        clabel(1) = ""
        ReDim vector(1 To UBound(.results.graphics), 1 To 2)
        For i = 1 To UBound(.results.graphics)
            vector(i, 2) = .results.graphics(i).strain_l_med / 1000000
            vector(i, 1) = .results.graphics(i).l_med * 1000000 ' micrometers
        Next i
        With frm_exp_data(doc).chart_exp_data.Plot.Axis(VtChAxisIdX)
            .ValueScale.Auto = False
            .AxisScale.LogBase = 10
            .AxisScale.Type = VtChScaleTypeLogarithmic
            .ValueScale.MinorDivision = 10
            .ValueScale.MajorDivision = 100
            
        End With
     End If
End With

With frm_exp_data(doc).chart_exp_data
    .RowCount = 1
    .ChartData = vector()
    .chartType = chart_type
    For i = 1 To num_rows
        .Row = i
        .RowLabelIndex = 1
        .RowLabel = ""
    Next i
    .ColumnCount = num_columns
    .ColumnLabelCount = num_columns
    For i = 1 To num_columns
        .Column = i
        .ColumnLabelIndex = 1
        .ColumnLabel = clabel(i)
    Next i
    
    .ShowLegend = True
    .Legend.Location.LocationType = VtChLocationTypeTop
    
    With .Plot.SeriesCollection(1)
        
        .DataPoints(-1).DataPointLabel.VtFont.name = "Verdana"
        .DataPoints(-1).DataPointLabel.VtFont.Size = 8
    End With
    With .Plot.SeriesCollection(2)
        .DataPoints(-1).DataPointLabel.VtFont.name = "Verdana"
        .DataPoints(-1).DataPointLabel.VtFont.Size = 8
    End With
    
    With .Plot.Axis(VtChAxisIdY)
        .AxisTitle = label_x
        .AxisScale.Hide = False
        .CategoryScale.Auto = True
    End With

    With .Plot.Axis(VtChAxisIdX)
        .AxisTitle = label_y
        .AxisScale.Hide = False
        .CategoryScale.Auto = True
        .AxisGrid.MajorPen.Style = VtPenStyleNull
        .AxisGrid.MinorPen.Style = VtPenStyleNull
    End With
    .Plot.UniformAxis = False
    .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
    .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
    .Visible = True
    .Refresh
End With

End Sub


Private Sub CopyGraph_Click()
Dim doc As Integer
doc = current_form

frm_exp_data(doc).chart_exp_data.EditCopy
End Sub

Private Sub exit_Click()
Dim doc As Integer
doc = current_form

frm_exp_data(doc).Hide
Unload Me
End Sub


Private Sub mnu_extension_Click()
frm_extension_selection.Show 1

End Sub

Private Sub printgraph_Click()
Dim doc As Integer
doc = current_form
'frm_exp_data(doc).PrintForm
 On Error GoTo vierror
  chht = chart_exp_data.Height
  chwd = chart_exp_data.Width
  chht1 = chht
  chwd1 = chwd
  Clipboard.Clear
  chart_exp_data.EditCopy
  Print_graph.Show 1
  Exit Sub
vierror:
 MsgBox Err.Description
End Sub

