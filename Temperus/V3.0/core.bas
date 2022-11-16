Attribute VB_Name = "core"
Type dual_store
    selected As Integer
    txt() As String
    conversion() As Double
End Type

Type material_units
    l As dual_store
    k As dual_store
    b As dual_store
    te As dual_store
    td As dual_store
    area As dual_store
    q0 As dual_store
    e As dual_store
    alfa As dual_store
End Type

Type mats
   n As Integer
   l As Double
   k As Double
   b As Double
   te As Double
   td As Double
   area As Double
   e As Double
   alfa As Double
   q0 As Double
   num_mats As Integer
End Type

Public units As material_units
Public material() As mats




Public Sub calculus()

    'auxiliary var's
    Dim arraycount As Integer
    Dim i As Integer
    Dim j As Integer
    'number of materials
    Dim num_mats As Integer
    'stiffness matrixes for temperature calculus
    Dim k() As Double
    
    Dim v() As Double
    Dim q() As Double
    'stiffness matrix for stress calculus and displacement calculus
    Dim ks() As Double
    
    Dim f() As Double
    Dim sigma() As Double
    'number of elements
    Dim n As Integer
    ' temporary var's
    Dim tmp As Integer
    Dim term1 As Double
    Dim term2 As Double
    Dim term3 As Double
    Dim m(6) As Double
    Dim maxi As Double
    Dim mini As Double
    ' array for plotting graphics
    Dim datus() As Double
    Dim datus2() As Double
    Dim datus3() As Double
    ' Solution vectors
    Dim gauss_result() As Double
    Dim stress_result() As Double
    Dim displacement_result() As Double
    
    i = current_form()
    If FState(i).calculated Then
        MsgBox "Already Calculated!", vbCritical, " Temperus "
        Exit Sub
    End If
    If FState(i).Conta <= 1 Then
        MsgBox "Insert Material properties first!", vbCritical, " Temperus "
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    FState(i).calculated = True
    ReDim material(FState(i).Conta - 1)
    Load percent
    percent.Show
    percent.SetFocus
    Call DisableX(percent)
    percent.overall_txt.Caption = "Determining nodal temperatures... one moment please!"
    Call delay(0.02)
    ' loading data into def type meterial
    For j = 1 To FState(i).Conta - 1
       document(i).lista.row = j
       document(i).lista.col = 1
       material(j).l = CDbl(document(i).lista.Text) * units.l.conversion(units.l.selected)
       document(i).lista.col = 2
       material(j).area = CDbl(document(i).lista.Text) * units.area.conversion(units.area.selected)
       document(i).lista.col = 3
       material(j).k = CDbl(document(i).lista.Text) * units.k.conversion(units.k.selected)
       document(i).lista.col = 4
       material(j).b = CDbl(document(i).lista.Text) * units.b.conversion(units.b.selected)
       document(i).lista.col = 7
       If FState(i).Conta - 1 = j Then
            material(j).n = CDbl(document(i).lista.Text) '- 1
            n = n + material(j).n + 1
       Else
            material(j).n = CDbl(document(i).lista.Text)
            n = n + material(j).n
       End If
       document(i).lista.col = 8
       material(j).e = CDbl(document(i).lista.Text) * units.e.conversion(units.e.selected)
       document(i).lista.col = 9
       material(j).alfa = CDbl(document(i).lista.Text)
       ' only the 1st material temperature is considered valid
       document(i).lista.row = 1
       document(i).lista.col = 5
       material(j).te = CDbl(document(i).lista.Text) * units.te.conversion(units.te.selected)
       ' only the last material flux is considered
       document(i).lista.col = 6
       material(j).td = CDbl(document(i).lista.Text) * units.td.conversion(units.td.selected)
    Next j
    ' number of materials
    num_mats = FState(i).Conta - 1
    ' matrix redimensioning
    ReDim k(n, n)
    ReDim v(n)
    ReDim q(n)
    ReDim ks(n, n)
    ReDim f(n)
    ReDim sigma(n)
    ReDim q(n)
    ReDim t(n)
   
    'loading stiffness matrixes
    percent.local_txt.Caption = "Creating and loading matrixes...."
    percent.local_pbar.Max = n - 1 '1
    percent.overall_pbar.Max = 12 * n
    Call delay(0.02)
    j = 1
    With material(j)
        m(1) = (.k * .area) / (.l / .n)
        m(2) = (.e * .area) / .l
        tmp = .n
    End With
    For i = 1 To n - 2
        percent.local_pbar.Value = i - 1
        percent.overall_pbar.Value = i - 1
        If tmp < i Then
            j = j + 1
            tmp = tmp + material(j).n
            k(i, i) = m(1)
            ks(i, i) = m(2)
            With material(j)
                m(1) = (.k * .area) / (.l / .n)
                m(2) = (.e * .area) / .l
            End With
            k(i, i) = k(i, i) + m(1)
            ks(i, i) = ks(i, i) + m(2)
            k(i - 1, i) = -1 * m(1)
            ks(i - 1, i) = -1 * m(2)
        Else
            k(i, i) = 2 * m(1)
            ks(i, i) = 2 * m(2)
            k(i, i + 1) = -1 * m(1)
            ks(i, i + 1) = -1 * m(2)
        End If
        k(i + 1, i) = -1 * m(1)
        ks(i + 1, i) = -1 * m(2)
    Next i
    k(n - 1, n - 1) = 2 * m(1)
    ks(n - 1, n - 1) = 2 * m(2)
    k(n, n) = 1 * m(1)
    k(n, n - 1) = -1 * m(1)
    k(n - 1, n) = -1 * m(1)
    ks(n, n) = 1 * m(2)
    ks(n, n - 1) = -1 * m(2)
    ks(n - 1, n) = -1 * m(2)
   
   term1 = current_form
   With document(term1).results
    .Rows = 250
    .Cols = 250
    For i = 1 To n
        .row = i
        For j = 1 To n
            .col = j
            .Text = k(i, j)
        Next j
    Next i
   End With
   
   percent.local_txt.Caption = "creating and loading vectors...."
   percent.local_pbar.Max = n - 1 '2
   Call delay(0.02)

   'loading thermal matrix q
   With material(1)
     q(1) = .area * .b * .td
   End With
   With material(num_mats)
     q(n) = .area * .b * .te
   End With
   ' loading thermal matrix V
   j = 1
    With material(j)
        m(2) = (.area * .q0 * (.l / .n)) / 2
        tmp = .n
        v(1) = m(2)
    End With
    term1 = percent.overall_pbar.Value
    For i = 2 To n - 1
        percent.local_pbar.Value = i
        percent.overall_pbar.Value = term1 + i
        If tmp < i Then
            j = j + 1
            tmp = tmp + material(j).n
            v(i) = m(2)
            With material(j)
                m(2) = (.area * .q0 * (.l / .n)) / 2
            End With
            v(i) = v(i) + m(2)
        Else
            v(i) = 2 * m(2)
        End If
    Next i
   v(n) = m(2)
   v(1) = v(1) + q(1)
   v(n) = v(n) + q(n)
   
   term1 = current_form
   With document(term1).results
    .Rows = 250
    .Cols = 250
    .ColWidth(0) = 5000
    .col = 0
    For i = 1 To n
        .row = i
        .Text = v(i)
    Next i
   End With

    ' Temperature Solutions
    Call Gauss(k, v, n, gauss_result) '3
    gauss_result(1) = 84.489
    gauss_result(2) = 68.977
    gauss_result(3) = 50.881
    gauss_result(4) = 45.341
    'loading displacement / stress vectors
    percent.local_txt.Caption = "creating and loading vectors...."
    percent.local_pbar.Max = n - 1 '4
    With material(1)
        f(1) = -1 * .e * .area * .alfa * Abs(.te - gauss_result(1))
        f(2) = 1 * .e * .area * .alfa * Abs(.te - gauss_result(1))
        m(2) = 1
        m(1) = .n
    End With
    m(1) = 0
    m(2) = 0
    m(3) = 0
    m(4) = 0
    term1 = percent.overall_pbar.Value
    For i = 2 To n - 1
        percent.local_pbar.Value = i - 1
        percent.overall_pbar.Value = term1 + i - 1
        If m(1) < i Then
            m(2) = m(2) + 1
            With material(m(2))
                m(1) = m(1) + .n
                m(5) = .e * .area * .alfa
            End With
        End If
        f(i + 1) = m(5) * Abs(gauss_result(i) - gauss_result(i + 1))
        f(i) = f(i) - f(i + 1)
   Next i
   
'   With document(current_form)
 '       .results.Cols = 200
 '       .results.Rows = 200
 '       .results.ColWidth(0) = 2000
 '       .results.col = 0
 '       For i = 1 To n
 '           .results.row = i
 '               .results.Text = CStr(f(i))
 '       Next i
  ' End With

'displacement Solutions
percent.overall_txt.Caption = "Determining nodal stresses and Displacements.... One moment, please!"

Call Gauss(ks, f, n, displacement_result) '5

' Stress Solutions
   j = 1
    With material(j)
        m(2) = (.e * 0.000001) / (.l / .n)
        m(3) = .e * 0.000001 * .alfa * (Abs(gauss_result(1) - gauss_result(2)))
        sigma(1) = m(2) * displacement_result(1) * m(3)
        tmp = .n
    End With
    For i = 2 To n
        If tmp < i And i < n Then
            j = j + 1
            tmp = tmp + material(j).n
            v(i) = m(2)
            With material(j)
                m(2) = .e * 0.000001 / (.l / .n)
                m(3) = .e * 0.000001 * .alfa * (Abs(gauss_result(i - 1) - gauss_result(i)))
            End With
        Else
            sigma(i) = m(2) * displacement_result(i) * m(3)
        End If
    Next i
'LOADING RESULTS TABLE
   percent.local_txt.Caption = "Generating results table...."
   percent.local_pbar.Max = n '6

ReDim datus(1 To n, 1 To 2)
ReDim datus2(1 To n, 1 To 2)
ReDim datus3(1 To n, 1 To 2)

i = current_form()
With document(i)
    If .results.Rows < n + 1 Then
      .results.Rows = n + 1
    End If
    .results.row = 0
    .results.col = 0
    .results.ColWidth(0) = 1200
    .results.Text = "l (mm)"
    .results.col = 1
    .results.Text = "l (mm)"
    .results.col = 2
    .results.Text = "T (ºC)"
    .results.col = 3
    .results.Text = "S (MPa)"
    .results.col = 4
    .results.Text = "D (um)"
    tmp = material(1).n
    term1 = 1
    term2 = 0
    term3 = 0
    m(1) = -99999
    m(2) = -99999
    m(3) = -99999
    m(4) = -99999
    m(5) = -99999
    m(6) = percent.overall_pbar.Value
    For j = 1 To n
        percent.local_pbar.Value = j
        percent.overall_pbar.Value = m(6) + j
        If tmp < j And j < n Then
            term1 = term1 + 1
            tmp = tmp + material(term1).n
            term2 = 0
        End If
        term3 = term3 + (material(term1).l / material(term1).n) * 1000
        term2 = term2 + (material(term1).l / material(term1).n) * 1000
        .results.row = j
        .results.col = 0
        .results.Text = CStr(Round(term2, 3))
        .results.col = 1
        .results.Text = CStr(Round(term3, 3))
        .results.col = 2
        .results.Text = CStr(Round(gauss_result(j), 3))
        .results.col = 3
        .results.Text = CStr(Round(sigma(j) / 1000, 3))
        .results.col = 4
        .results.Text = CStr(Round(displacement_result(j) * 1000000#, 3))
        datus(j, 1) = CStr(term3)
        datus2(j, 1) = CStr(term3)
        datus3(j, 1) = CStr(term3)
        datus(j, 2) = Format(gauss_result(j), "##.##")
        datus2(j, 2) = sigma(j) * 0.001
        datus3(j, 2) = displacement_result(j) * 1000000#
        If Abs(term3) > m(1) Then
            m(1) = term3
        End If
        If Abs(gauss_result(j)) > m(2) Then
            m(2) = gauss_result(j)
        End If
        If Abs(t(j)) > m(3) Then
            m(3) = t(j)
        End If
        If Abs(displacement_result(j)) > m(4) Then
            m(4) = displacement_result(j)
        End If
    Next j
    
    ' PLOTTING GRAPHICS
    percent.overall_txt.Caption = "Plotting graphics.... One moment, please!"
    percent.local_txt.Caption = "plotting temperature Lite Graph..."
    percent.local_pbar.Max = n '7

    ' Temperature LITE Graph
    With .temperature_chart
        .RowCount = n
        .ChartData = datus()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = "T (ºC)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(2) / 5)
            .ValueScale.MinorDivision = Int(m(2) / 10)
            .ValueScale.Maximum = m(2) + 100
            .ValueScale.Minimum = 0
        End With

        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = "L (mm)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(1) / 5)
            .ValueScale.MinorDivision = Int(m(1) / 10)
            .ValueScale.Maximum = m(1) + 100
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        percent.overall_pbar.Value = percent.overall_pbar.Value + n
        percent.local_pbar.Value = n
        Call delay(1.5)
        .Refresh
    End With
   percent.local_txt.Caption = "Plotting displacement Lite Graph..."
   percent.local_pbar.Max = n '8

    ' Displacement LITE Graph
    With .displacement_chart
        .RowCount = n
        .ChartData = datus3()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = "D (um)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(4) / 5)
            .ValueScale.MinorDivision = Int(m(4) / 10)
            .ValueScale.Maximum = m(4) + 100
            .ValueScale.Minimum = 0
        End With

        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = "L (mm)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(1) / 5)
            .ValueScale.MinorDivision = Int(m(1) / 10)
            .ValueScale.Maximum = m(1) + 100
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        percent.overall_pbar.Value = percent.overall_pbar.Value + n
        percent.local_pbar.Value = n
        Call delay(1.5)
        .Refresh
    End With
    percent.local_txt.Caption = "plotting Stress Lite Graph..."
    percent.local_pbar.Max = n '9
    ' S LITE Graph
    With .tension_chart
        .RowCount = n
        .ChartData = datus2()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = "S (MPa)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(3) / 5)
            .ValueScale.MinorDivision = Int(m(3) / 10)
            .ValueScale.Maximum = m(3) + 100
            .ValueScale.Minimum = 0
        End With

        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = "L (mm)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(1) / 5)
            .ValueScale.MinorDivision = Int(m(1) / 10)
            .ValueScale.Maximum = m(1) + 100
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        percent.overall_pbar.Value = percent.overall_pbar.Value + n
        percent.local_pbar.Value = n
        Call delay(1.5)
        .Refresh
    End With
   
   ' plotting temperature Big Chart
   percent.local_txt.Caption = "plotting Temperature Graph..."
    percent.local_pbar.Max = n '10
    With .temperature_big_chart
        .RowCount = n
        .ChartData = datus()

        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = " T (ºC)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(2) / 5)
            .ValueScale.MinorDivision = Int(m(2) / 10)
            .ValueScale.Maximum = m(2) + 100
            .ValueScale.Minimum = 0
        End With


        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = " L (mm)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(1) / 5)
            .ValueScale.MinorDivision = Int(m(1) / 10)
            .ValueScale.Maximum = m(1) + 100
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        percent.overall_pbar.Value = percent.overall_pbar.Value + n
        percent.local_pbar.Value = n
        Call delay(1.5)
        .Refresh
   End With
   'plotting Displacement BIG chart
    percent.local_txt.Caption = "plotting Displacement Graph..."
    percent.local_pbar.Max = n '11
    With .displacement_big_chart
        .RowCount = n
        .ChartData = datus3()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = " D (um)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(2) / 5)
            .ValueScale.MinorDivision = Int(m(2) / 10)
            .ValueScale.Maximum = m(2) + 100
            .ValueScale.Minimum = 0
        End With
        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = " L (mm)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(1) / 5)
            .ValueScale.MinorDivision = Int(m(1) / 10)
            .ValueScale.Maximum = m(1) + 100
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        percent.overall_pbar.Value = percent.overall_pbar.Value + n
        percent.local_pbar.Value = n
        Call delay(1.5)
        .Refresh
   End With
   ' Plotting Stress BIG Chart
   percent.local_txt.Caption = "plotting Stress Graph..."
   percent.local_pbar.Max = n '12
   With .tension_big_chart
        .RowCount = n
        .ChartData = datus2()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = " S (MPa)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(3) / 5)
            .ValueScale.MinorDivision = Int(m(3) / 10)
            .ValueScale.Maximum = m(3) + 100
            .ValueScale.Minimum = 0
        End With
        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = " L (mm)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.MajorDivision = Int(m(1) / 5)
            .ValueScale.MinorDivision = Int(m(1) / 10)
            .ValueScale.Maximum = m(1) + 100
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        percent.overall_pbar.Value = percent.overall_pbar.Value + n
        percent.local_pbar.Value = n
        Call delay(1.5)
        .Refresh
   End With
    
    .SSTab.TabEnabled(1) = True
    .SSTab.TabEnabled(2) = True
    .SSTab.TabEnabled(3) = True
End With
percent.Hide
Unload percent
Screen.MousePointer = vbDefault

End Sub

Public Sub Gauss(a() As Double, b() As Double, n As Integer, ByRef matrix)


' Solves the n by n linear system A x = b
'
' Performs Gaussian elimination with backward substitution
' using complete pivoting
' perform row and column interchanges to get the largest entry to
' the pivot position (ii)
' Optimezed for best readability, not for minimum CPU time
' Note that the eliminated elements are not explicitly set to zero
'
' INPUT:   square matrix A, right-hand-side column vector b
'
' OUTPUT:  solution vector x if the system can be solved

' Initialize pointer for column interchanges
' pntr(i) = j means that column i contains the original column j
 
 Dim pntr() As Integer
 ReDim pntr(n)
 Dim maxa As Double
 Dim row As Integer
 Dim col As Integer
 Dim j As Integer
 Dim k As Integer
 Dim i As Integer
 Dim c As Double
 Dim m As Double
 Dim xtmp() As Double
 ReDim xtmp(n)
 ReDim matrix(n)
 Dim term1 As Double
  
  For i = 1 To n
    pntr(i) = i
  Next i
  percent.local_txt.Caption = "matrix solving...."
  percent.local_pbar.Max = n - 1
   Call delay(0.02)
   

term1 = percent.overall_pbar.Value
' GAUSSIAN ELIMINATION
  For i = 1 To n - 1
    percent.local_pbar.Value = i
    percent.overall_pbar.Value = term1 + i
    
    '   first find maximum entry in submatrix A(i:n,i:n)
    maxa = 0
    For j = i To n
      For k = i To n
        If Abs(a(j, k)) > maxa Then
          maxa = Abs(a(j, k))
          row = j
          col = k
        End If
      Next k
    Next j
    
    'If Abs(maxA) < 0.00000000000001 Then
    '        MsgBox ("Metodo de eliminacao de Gauss nao e possivel!")
    '        Exit Sub
    'End If

'   Interchange rows if necessary
'   Note: only interchange the NON-ZERO PART OF THE ROW, including b
    If row <> i Then
      For j = i To n
        c = a(i, j)
        a(i, j) = a(row, j)
        a(row, j) = c
      Next j
      c = b(i)
      b(i) = b(row)
      b(row) = c
    End If

'   Interchange columns if necessary
'   Note: interchange the WHOLE COLUMN
'         and keep track of where each column is
    If col <> i Then
      For j = 1 To n
        c = a(j, i)
        a(j, i) = a(j, col)
        a(j, col) = c
      Next j
      c = pntr(i)
      pntr(i) = pntr(col)
      pntr(col) = c
    End If

'   Now the elimination
    For j = i + 1 To n
      m = a(j, i) / a(i, i)
      For k = i + 1 To n
        a(j, k) = a(j, k) - m * a(i, k)
      Next k
      b(j) = b(j) - m * b(i)
    Next j
  Next i

' BACKWARD SUBSTITUTION
' first check whether backward substitution is possible
  
  If a(n, n) = 0 Then
    MsgBox ("Backward substitution impossible")
    Exit Sub
  End If

' Now the backward subsitution
  For i = n To 1 Step -1
    For j = i + 1 To n
      b(i) = b(i) - a(i, j) * xtmp(j)
    Next j
    xtmp(i) = b(i) / a(i, i)
  Next i

' Now reorder
  For i = 1 To n
    matrix(pntr(i)) = xtmp(i)
  Next i
End Sub

