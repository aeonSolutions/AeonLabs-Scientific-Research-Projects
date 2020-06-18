Attribute VB_Name = "core"
Type listas
    tf As String
    l As String
    ef As String
    vf As String
    ts As String
    vs As String
    es As String
    sigma As String
    delta As String
    rs As String
    m As String
    sigma_weib As String
    n As String
    divisoes As String
End Type

Type tabela
    l As Boolean
    init As Boolean
    final As Boolean
    blocks As Boolean
    sigma As Boolean
    sigmaf As Boolean
    coord As Boolean
    cracked As Boolean
    pos(8) As Integer
End Type
Type elemental
    sigma As Double
    init As Double
    final As Double
    blocks As Double
    coord As Double
    cracked As Boolean
    counter As Integer
    sigmaf As Double
End Type

Public lista As listas
Public colunas As tabela
Public elements() As elemental

Public k As Integer
Public j As Integer
Public i As Integer
Public beta As Double
Public epslon As Double
Public vs As Double
Public vf As Double
Public es As Double
Public ef As Double
Public tf As Double
Public ts As Double
Public num_elements As Integer
Public num_voltas As Integer
Public sigma As Double
Public delta_sigma As Double
Public l As Double
Public sigma_weib As Double
Public rs As Double
Public sigma_p As Double
Public num_cracks As Integer
Public tmp As Double
Public m As Double
Public contador As Integer
Public D As Double
Public stats() As Double
Public maxx As Double
Public stats2() As Double
Public maxx2 As Double
Public first_time As Boolean
Public num_segments As Integer
Public generated As Boolean
Public stats3() As Double






Public Function sinh(x As Double) As Double
    On Error Resume Next
    If Err <> 0 Then
        If x < 0 Then
            sinh = -100
        Else
            sinh = 100
        End If
    Else
        sinh = (Exp(x) - Exp(-x)) / 2
    End If
End Function
Public Function cosh(x As Double) As Double
    On Error Resume Next
    If Err <> 0 Then
        If x < 0 Then
            cosh = -100
        Else
            cosh = 100
        End If
    Else
        cosh = (Exp(x) + Exp(-x)) / 2
    End If
End Function
Public Function tanh(x As Double) As Double
    Dim test As Double
    
    On Error Resume Next
    test = Exp(x)
    If Err <> 0 Then
        If x < 0 Then
            tanh = -1
        Else
            tanh = 1
        End If
    Else
        tanh = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
    End If
End Function

Sub stress_load()
Dim tmp3 As Double
Dim tmp2 As Double
Dim curr As Integer

 curr = current_form
 With document(curr)
  If Not IsNumeric(.sigma_txt.Text) Then
    .sigma_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.m_txt.Text) Then
    .m_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.rs_txt.Text) Then
    .rs_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.sigma_weib_txt.Text) Then
    .sigma_weib_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.divisoes_txt.Text) Then
    .divisoes_txt.SetFocus
    Exit Sub
  End If
  sigma_weib = CDbl(.sigma_weib_txt) * 1000000#
  sigma = CDbl(.sigma_txt) * 1000000#
  sigma = sigma + rs
  m = CDbl(.m_txt)
  num_elements = CDbl(.divisoes_txt.Text)
End With
On Error Resume Next
tmp2 = UBound(elements)
If Err <> 0 Then
    ReDim elements(num_elements)
    Err.Clear
End If
 tmp2 = 9E+40
 For i = 1 To num_elements
    Randomize
    tmp3 = Rnd
    elements(i).sigmaf = sigma_weib * (-Log(tmp3)) ^ (1 / m)
    If elements(i).sigmaf < tmp2 Then
        tmp2 = elements(i).sigmaf
    End If
    elements(i).init = 1
    elements(i).final = num_elements
    elements(i).blocks = num_elements
    elements(i).coord = 0
    elements(i).cracked = False
    elements(i).counter = 1
    elements(i).sigma = sigma - delta_sigma
 Next i
generated = True
document(curr).lowest_txt.Text = Str(Round(tmp2 / 1000000#, 2))
MsgBox "Probabilistic Stresses Generated Sucssefully!", vbOKCancel, "Info"

End Sub

Sub check_crack()
    Dim tmp As Double
    Dim cur_doc As Integer

num_cracks = 0
percent.local_txt.Caption = "verifying existing cracks..."


For k = 1 To num_elements

    If elements(k).sigmaf < elements(k).sigma Then
        elements(k).cracked = True
        elements(k).sigma = 0
        num_cracks = num_cracks + 1
    End If
Next k
If first_time And num_cracks > 0 Then

    cur_doc = current_form
    tmp = 9E+40
    For k = 1 To num_elements
        If elements(k).cracked And elements(k).sigmaf < tmp Then
            tmp = elements(k).sigmaf
        End If
    Next k
    document(cur_doc).sf_txt.Text = Str(Round(tmp / 1000000#, 2))
    first_time = False
End If

End Sub
Sub new_elements()
  Dim pos() As Integer
  ReDim pos(num_elements, 2)
  Dim last_crack As Boolean
  
 percent.local_txt.Caption = "creating Sub-elements... [" & Str(num_cracks) & "]"
 Call delay(0.01)
 contador = 0
 last_crack = True
 For i = 1 To num_elements

    If elements(i).cracked = False Then
        If last_crack = True Then
            contador = contador + 1
            pos(contador, 1) = i
            last_crack = False
        End If
        pos(contador, 2) = pos(contador, 2) + 1
    Else
        last_crack = True
    End If
 Next i
For i = 1 To contador
    k = 0
    For j = pos(i, 1) To pos(i, 1) + pos(i, 2) - 1
        elements(j).init = pos(i, 1)
        elements(j).final = pos(i, 1) + pos(i, 2) - 1
        elements(j).blocks = pos(i, 2)
        elements(j).coord = k * D
        elements(j).counter = k + 1
        k = k + 1
    Next j
 Next i
End Sub

Sub new_stress()
percent.local_txt.Caption = "Solving for Stresses..."
Call delay(0.01)
For i = 1 To num_elements

    If elements(i).cracked = False Then
        epslon = elements(i).blocks * D / 2
        elements(i).sigma = elements(i).sigma * (tanh(beta * epslon) * sinh(beta * (elements(i).coord)) - cosh(beta * (elements(i).coord)) + 1)
    If elements(i).sigma < 0 Then
        MsgBox "Erro!: Tensao negativa [" & Str(Round(elements(i).sigma, 2)) & "] no elemento nº" & Str(i), vbOKCancel, "Info"
    End If
    End If
Next i

End Sub
Sub engine()
Dim cur_doc As Integer

cur_doc = current_form
percent.overall_txt.Caption = "Solving Cracking... One moment, please!"
percent.overall_pbar.Max = num_voltas + 1
Call delay(0.02)

first_time = True
num_cracks = -1
With document(cur_doc)
sigma = sigma - delta_sigma
For ii = 1 To num_voltas + 1
    sigma = sigma + delta_sigma
    percent.overall_pbar.Value = ii
    Call delay(0.01)
    For k = 1 To num_elements
        elements(k).sigma = elements(k).sigma + delta_sigma
    Next
    If num_cracks = 0 Then
        num_cracks = -1
    End If
    While num_cracks <> 0
        Call check_crack
        If num_cracks <> 0 Then
            Call delay(0.03)
            Call new_elements
            Call new_stress
        End If
    Wend
Next ii

End With
Call delay(0.5)
End Sub

Public Sub Calcular()
Dim cur_doc As Integer
Dim tmp2 As Integer



cur_doc = current_form
With document(cur_doc)
  If Not IsNumeric(.tf_txt.Text) Then
    .tf_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.l_txt.Text) Then
    .l_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.ef_txt.Text) Then
    .ef_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.vf_txt.Text) Then
    .vf_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.ts_txt.Text) Then
    .ts_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.es_txt.Text) Then
    .es_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.vs_txt.Text) Then
    .vs_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.sigma_txt.Text) Then
    .sigma_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.delta_txt.Text) Then
    .delta_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.rs_txt.Text) Then
    .rs_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.m_txt.Text) Then
    .m_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.sigma_weib_txt.Text) Then
    .sigma_weib_txt.SetFocus
    Exit Sub
    End If
  If Not IsNumeric(.n_txt.Text) Then
    .n_txt.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(.divisoes_txt.Text) Then
    .divisoes_txt.SetFocus
    Exit Sub
  End If
  If Not IsNumeric(.segments_txt.Text) Then
    .segments_txt.SetFocus
    Exit Sub
  End If
  num_segments = CDbl(.segments_txt.Text)
  num_voltas = CDbl(.n_txt.Text)
  ef = CDbl(.ef_txt) * 1000000000#
  es = CDbl(.es_txt) * 1000000000#
  tf = CDbl(.tf_txt) * 0.001
  ts = CDbl(.ts_txt) * 0.001
  l = CDbl(.l_txt) * 0.01
  sigma_weib = CDbl(.sigma_weib_txt) * 1000000#
  sigma = CDbl(.sigma_txt) * 1000000#
  rs = CDbl(.rs_txt) * 1000000#
  sigma = sigma + rs
  delta_sigma = CDbl(.delta_txt) * 1000000#
  m = CDbl(.m_txt)
  vs = CDbl(.vs_txt.Text)
  vf = CDbl(.vf_txt)
  beta = Sqr((1 - vf) / tf ^ 2 + (ef * (1 - vs) ^ 2) / (es * ts * tf * (1 + vf)))
  .beta_txt.Text = CStr(Round(beta, 8))
If generated = False Then
    Call stress_load
Else
    For i = 1 To num_elements
       elements(i).init = 1
       elements(i).final = num_elements
       elements(i).blocks = num_elements
       elements(i).coord = 0
       elements(i).cracked = False
       elements(i).counter = 1
       elements(i).sigma = sigma
    Next i
End If
  D = l / num_elements
  Load percent
  percent.Show
  percent.SetFocus
  Call DisableX(percent)
  Call delay(0.02)
  Call engine
  .curr_stress_txt.Text = Str(sigma / 1000000#)
  num_cracks = 0
  For i = 1 To num_elements
    If elements(i).cracked Then
        num_cracks = num_cracks + 1
    End If
  Next i
  .num_cracks_txt.Text = Str(num_cracks)
  tmp = num_elements / num_segments
  ReDim stats(1 To num_segments, 1)
  tmp2 = tmp
  contador = 1
  For i = 1 To num_elements
    If i = tmp2 + 1 Then
        contador = contador + 1
        tmp2 = tmp2 + tmp
    End If
    If elements(i).cracked Then
        stats(contador, 1) = stats(contador, 1) + 1
        elements(i).init = i
        elements(i).final = i
        elements(i).blocks = 1
        elements(i).coord = 0
        
    End If
  Next i
  document(cur_doc).elements_per_segment_txt.Caption = Str(Round(tmp, 2)) & " elements per segment"
  maxx2 = -999
  tmp2 = 0
  contador = 0
  ' Generating the stresses data table
  ReDim stats2(1 To num_elements, 1 To 2)
  For i = 1 To num_elements
    If elements(i).cracked Then
        stats2(i, 2) = 0
        tmp2 = tmp2 + 1
    Else
        If contador < tmp2 Then
            contador = tmp2
        End If
        tmp2 = 0
        stats2(i, 2) = CDbl(elements(i).sigma / 1000000#)
    End If
    If stats2(i, 2) > maxx2 Then
        maxx2 = stats2(i, 2)
    End If
    stats2(i, 1) = i
  Next i
  ' Generating the frequencies data table
  If contador < 8 Then
    contador = 8
  End If
  ReDim stats3(1 To contador, 1)
  'For i = 1 To contador
 '   stats3(i, 1) = i
 ' Next i
  contador = 0
  For i = 1 To num_elements
  If elements(i).cracked Then
      contador = contador + 1
  Else
    If contador > 0 Then
        stats3(contador, 1) = stats3(contador, 1) + 1
    End If
    contador = 0
  End If
  
  Next i
  percent.Hide
  Unload percent
  Call load_results

End With
FState(cur_doc).calculated = True
End Sub

Sub load_results()
Dim cur_doc As Integer

cur_doc = current_form

With document(cur_doc)
     If num_elements = 0 Then
       Exit Sub
     End If
     Load percent
     percent.Show
     percent.SetFocus
     percent.overall_txt.Caption = "Loading data..."
     Call delay(0.02)

     percent.overall_pbar.Max = num_elements
     .results.Rows = num_elements + 1
     With colunas
        .l = True
        .blocks = True
        .coord = True
        .cracked = True
        .final = True
        .init = True
        .sigma = True
        .sigmaf = True
        .pos(5) = 1
        .pos(6) = 2
        .pos(8) = 3
        .pos(7) = 4
        .pos(1) = 5
        .pos(2) = 6
        .pos(3) = 7
        .pos(4) = 8
     End With
     .results.Visible = False
     .results.Clear
     Call colwidth
    For i = 1 To num_elements
        .results.Row = i
        .results.Col = 0
        .results.Text = Str(i)

        percent.overall_pbar.Value = i
        If colunas.l Then
            .results.Col = colunas.pos(1)
            .results.CellAlignment = 4
            .results.Text = CStr(Round(elements(i).blocks * D * 1000, 8))
        End If
        If colunas.init Then
            .results.Col = colunas.pos(2)
            .results.CellAlignment = 4
            .results.Text = CStr(elements(i).init)
        End If
        If colunas.final Then
            .results.Col = colunas.pos(3)
            .results.CellAlignment = 4
            .results.Text = Str(elements(i).final)
        End If
        If colunas.blocks Then
            .results.Col = colunas.pos(4)
            .results.CellAlignment = 4
            .results.Text = CStr(elements(i).blocks)
        End If
        If colunas.coord Then
            .results.Col = colunas.pos(7)
            .results.CellAlignment = 4
            .results.Text = CStr(Round(elements(i).coord * 1000, 8))
        End If
        If colunas.cracked Then
            .results.Col = colunas.pos(8)
            .results.CellAlignment = 4
            .results.Text = CStr(elements(i).cracked)
        End If
        If colunas.sigma Then
            .results.Col = colunas.pos(5)
            .results.CellAlignment = 4
            .results.Text = CStr(Round(elements(i).sigma / 1000000#, 5))
        End If
        If colunas.sigmaf Then
            .results.Col = colunas.pos(6)
            .results.CellAlignment = 4
            .results.Text = CStr(Round(elements(i).sigmaf / 1000000#, 5))
        End If
    Next i
    If num_cracks > 0 Then

        .SSTab.TabEnabled(2) = True
        ' Chart
        
        With .chart2
            .RowCount = num_elements
            .ColumnCount = 2
            .ChartData = stats2()
            .chartType = VtChChartType2dXY
            ' Set Guide Lines
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
            With .Plot.Axis(VtChAxisIdY)
                .AxisTitle = "Stress (MPa)"
                .AxisScale.Hide = False
                .CategoryScale.Auto = True
            End With
            With .Plot.Axis(VtChAxisIdX)
                .AxisTitle = "element nº"
                .AxisScale.Hide = False
                .CategoryScale.Auto = True
            End With
            .Plot.UniformAxis = False
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
            .Visible = True
            .Refresh
        End With
    
        .SSTab.TabEnabled(1) = True
        ' Chart

        With .chart
            .ChartData = stats()
            .chartType = VtChChartType2dBar
            ' Set Guide Lines
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
            With .Plot.Axis(VtChAxisIdY)
                .AxisTitle = "nº cracks"
                .AxisScale.Hide = False
                .CategoryScale.Auto = True
                .ValueScale.MinorDivision = 1
            End With
            With .Plot.Axis(VtChAxisIdX)
                .AxisTitle = "Segment nº"
                .AxisScale.Hide = False
                .CategoryScale.Auto = True
                .ValueScale.MinorDivision = 1
            End With
            .Plot.UniformAxis = False
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
            .Visible = True
            .Refresh
        End With
        .SSTab.TabEnabled(3) = True
        ' Chart

        With .chart3
            .ChartData = stats3()
            .chartType = VtChChartType2dBar
            ' Set Guide Lines
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
            With .Plot.Axis(VtChAxisIdY)
                .AxisTitle = "nº cracks"
                .AxisScale.Hide = False
                .CategoryScale.Auto = True
                .ValueScale.MinorDivision = 1
            End With
            With .Plot.Axis(VtChAxisIdX)
                .AxisTitle = "nº consecutive cracks"
                .AxisScale.Hide = False
                .CategoryScale.Auto = True
                .ValueScale.MinorDivision = 1
            End With
            .Plot.UniformAxis = False
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
            .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
            .Visible = True
            .Refresh
        End With
    
    End If
     .results.Visible = True
    percent.Hide
    Unload percent
End With

End Sub

Sub colwidth()
Dim cur_doc As Integer

cur_doc = current_form
With document(cur_doc)
    .results.Row = 0
    .results.Col = 0
    .results.MergeCol(5) = True
    .results.MergeCol(6) = True
    .results.MergeCol(7) = True
    .results.MergeCol(8) = True
    .results.Refresh
    .results.Text = "nº"
    .results.colwidth(1) = .TextWidth("#########") * 2
    .results.Col = 1
    .results.CellAlignment = 4
    .results.Text = "current stress(MPa)"
    .results.colwidth(2) = .TextWidth("########") * 2
    .results.Col = 2
    .results.CellAlignment = 4
    .results.Text = "Final stress(MPa)"
    .results.colwidth(3) = .TextWidth("######") * 2
    .results.Col = 3
    .results.CellAlignment = 4
    .results.Text = "Cracked ?"
    .results.colwidth(4) = .TextWidth("######") * 2
    .results.Col = 4
    .results.CellAlignment = 4
    .results.Text = "Coord (mm)"
    .results.colwidth(5) = .TextWidth("######") * 2
    .results.Col = 5
    .results.CellAlignment = 4
    .results.Text = "L (mm)"
    .results.colwidth(6) = .TextWidth("######") * 2
    .results.Col = 6
    .results.CellAlignment = 4
    .results.Text = "initial pos"
    .results.colwidth(7) = .TextWidth("#####") * 2
    .results.Col = 7
    .results.CellAlignment = 4
    .results.Text = "Final pos"
    .results.colwidth(8) = .TextWidth("#####") * 2
    .results.Col = 8
    .results.CellAlignment = 4
    .results.Text = "nº blocks"
End With
End Sub

