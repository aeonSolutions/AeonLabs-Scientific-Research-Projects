Attribute VB_Name = "Engine"
Option Explicit 'Transformas all declared variables into global variables

Private first_time As Boolean
Private redim_time As Boolean

Private d As Double
Private beta As Double
Private modulus() As Double


'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Function for generating probabilistic stresses
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Sub stress_load()
Dim tmp, tmp2, tmp3 As Double
Dim doc As Integer
Dim i As Integer

doc = current_form
With doc_props(doc)
    d = .phisical.lenght / .statistic.elements
    ReDim .elements(.statistic.elements)
    tmp2 = 9E+40
    tmp = 9E-40
    For i = 1 To .statistic.elements
        Randomize
        tmp3 = Rnd
        .elements(i).sigmaf = .statistic.sl + .statistic.s0 * (.statistic.elements * Log(1 / (1 - tmp3))) ^ (1 / .statistic.m)
        If .elements(i).sigmaf < tmp2 Then
            tmp2 = .elements(i).sigmaf
        End If
        If .elements(i).sigmaf > tmp Then
            tmp = .elements(i).sigmaf
        End If
        .elements(i).init = 1
        .elements(i).final = .statistic.elements
        .elements(i).blocks = .statistic.elements
        .elements(i).coord = (i - 1) * d
        .elements(i).cracked = False
        .elements(i).counter = 1
        .elements(i).sigma = 0
    Next i
    .elements_generated = True
    .results.lowest_rnd = Round(tmp2, 2) ' in Pa
    .results.highest_rnd = Round(tmp, 2) ' in Pa
End With

MsgBox "Probabilistic Stresses Generated Sucssefully!", vbOKCancel, "Info"
Call refresh_richtext
End Sub

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Function for checking the existence of cracks
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Function check_crack(cur_doc As Integer, strain As Double) As Integer

    Dim tmp As Double
    Dim k As Integer
    Dim num_cracks As Integer
    
num_cracks = 0
'frm_perform_calculations.txt.Caption = "verifying existing cracks..."

With doc_props(cur_doc)
    For k = 1 To .statistic.elements
        If .elements(k).sigmaf < .elements(k).sigma Then
            .elements(k).cracked = True
            .elements(k).sigma = 0
            num_cracks = num_cracks + 1
        End If
    Next k
    If first_time And num_cracks > 0 Then
        .results.crack_strain = strain
        tmp = 9E+40
        For k = 1 To .statistic.elements
            If .elements(k).cracked Then
                If .elements(k).sigmaf < tmp Then
                    tmp = .elements(k).sigmaf
                End If
            End If
        Next k
        .results.sf = tmp
        first_time = False
    End If
    num_cracks = 0
    For k = 1 To .statistic.elements
        If .elements(k).cracked Then
                num_cracks = num_cracks + 1
        End If
    Next k
End With

If num_cracks > 0 Then
    doc_props(cur_doc).results.cracks = num_cracks
    check_crack = num_cracks
    Exit Function
End If

End Function
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Function for generating new elements after cracking
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Sub new_elements(doc As Integer)
  Dim pos() As Integer
  ReDim pos(doc_props(doc).statistic.elements, 2)
  Dim last_crack As Boolean
  Dim contador, i, k, j As Integer
  
 'frm_perform_calculations.txt.Caption = "creating Sub-elements... [" & Str(num_cracks) & "]"
 contador = 0
 last_crack = True
 With doc_props(doc)
     For i = 1 To .statistic.elements
        If .elements(i).cracked = False Then
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
            .elements(j).init = pos(i, 1)
            .elements(j).final = pos(i, 1) + pos(i, 2) - 1
            .elements(j).blocks = pos(i, 2)
            .elements(j).coord = k * d
            .elements(j).counter = k + 1
            k = k + 1
        Next j
     Next i
End With
End Sub

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Function for generating new stresses after cracking
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Sub new_stress(doc As Integer, extension As Double)

Dim i As Integer
Dim epslon As Double
Dim l_block As Double
Dim b As Double
Dim ds0 As Double ' delta sigma zero
Dim e2l As Double ' epslon 2 linhas
Dim part1 As Double
Dim part2 As Double
Dim part3 As Double
Dim part4 As Double

'frm_perform_calculations.txt.Caption = "Solving for Stresses..."
With doc_props(doc)
    If .results.cracks = 0 Then
        For i = 1 To .statistic.elements
                .elements(i).sigma = .statistic.rs + .phisical.ef * (extension + .phisical.efr)
        Next i
        ds0 = 1
    Else ' o revestimento ja se encontra fissurado
        i = 0
        For i = 1 To .statistic.elements
            l_block = .elements(i).blocks * .phisical.lenght / .statistic.elements
            With .phisical
                ds0 = (.ts + .tf) / .ts * doc_props(doc).elements(i).sigma_exp - .es * (extension + .esr)
                e2l = extension + .efr - 1 / .ef * .ts / .tf * ds0
                b = .ts * (ds0 + ((.es * .ef * .tf * e2l) / (.ef * .tf + .es * .ts))) * (1 + Exp(-l_block / doc_props(doc).results.lambda)) ^ (-1)
            End With
            If .elements(i).cracked = False Then
                part1 = b / .phisical.tf
                part2 = 1 + Exp(-l_block / .results.lambda)
                part3 = Exp(-.elements(i).coord / .results.lambda) + Exp((.elements(i).coord - l_block) / .results.lambda)
                .elements(i).sigma = part1 * (part2 - part3)
            End If
        Next i
    End If
End With
End Sub

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Function for generating stresses for the Lmed graphics
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Sub new_stress_l_med(doc As Integer, extension As Double)

Dim i As Integer
Dim k As Integer
Dim epslon As Double
Dim l_block As Double
Dim b As Double
Dim ds0 As Double ' delta sigma zero
Dim e2l As Double ' epslon 2 linhas
Dim part1 As Double
Dim part2 As Double
Dim part3 As Double
Dim part4 As Double
Dim maxx As Double

With doc_props(doc)
    extension = extension / 100
    If .results.cracks = 0 Then
        ' put a flag in the vars
        doc_props(doc).results.graphics(extension * 100).strain_l_med = .elements(1).sigma_exp
        doc_props(doc).results.graphics(extension * 100).flag = False
    Else ' o revestimento ja se encontra fissurado
        maxx = -1
        ' procurar o comprimento interfissuras mais comprido
        For i = 1 To .statistic.elements
            If .elements(i).blocks > maxx And .elements(i).cracked = False Then
                k = i
                maxx = .elements(i).blocks
            End If
        Next i
        
        'determinar o comprimento max em metros
        l_block = maxx * .phisical.lenght / .statistic.elements ' Lmax
        'l_block = .phisical.lenght / doc_props(doc).results.cracks ' Lmed
        
        If redim_time Then
            redim_time = False
            doc_props(doc).results.graphics(1).max_elements = Round(l_block * .statistic.elements / .phisical.lenght, 0)
        End If
        
        doc_props(doc).results.graphics(extension * 100).l_med = l_block
        doc_props(doc).results.graphics(extension * 100).crk_density = 1 / l_block
        doc_props(doc).results.graphics(extension * 100).flag = True
        doc_props(doc).results.graphics(extension * 100).strain = extension * 100
        
        'plotar o grafico nao homogenizado
        For i = 1 To .statistic.elements
            .results.graphics(extension * 100).l(i) = .phisical.lenght / .statistic.elements * i
            .results.graphics(extension * 100).sf(i) = 0
        Next i
        For i = .elements(k).init To .elements(k).final
            .results.graphics(extension * 100).sf(i - .elements(k).init + 1) = .elements(i).sigma
            part4 = Round((.elements(k).final + .elements(k).init) / 2, 0)
            If part4 = i Then
                doc_props(doc).results.graphics(extension * 100).strain_l_med = .elements(i).sigma
            End If
        Next i
        
        'plotar i grafico homogenizado a Li/100 ptos
        part4 = l_block / 100
        For i = 1 To 100
            With .phisical
                ds0 = (.ts + .tf) / .ts * doc_props(doc).elements(k).sigma_exp - .es * (extension + .esr)
                e2l = extension + .efr - 1 / .ef * .ts / .tf * ds0
                b = .ts * (ds0 + ((.es * .ef * .tf * e2l) / (.ef * .tf + .es * .ts))) * (1 + Exp(-l_block / doc_props(doc).results.lambda)) ^ (-1)
            End With
            part2 = 1 + Exp(-l_block / .results.lambda)
            part3 = Exp(-part4 / .results.lambda) + Exp((part4 * (i - 1) - l_block) / .results.lambda)
            .results.graphics(extension * 100).hsf(i) = b / .phisical.tf * (part2 - part3)
            .results.graphics(extension * 100).hl(i) = part1 * i
        Next i
    End If
End With
End Sub
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Main Function for cracking analysis
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Sub engine()
Dim cur_doc As Integer
Dim k As Integer
Dim sigma, sig_now, sig_last As Double
Dim es As Double
Dim num_cracks, last_cracks As Integer
Dim factor As Double
Dim delta As Double
Dim ii As Integer

cur_doc = current_form
frm_perform_calculations.txt.Caption = "Solving Cracking... One moment, please!"
frm_perform_calculations.pbar.Max = Round(doc_props(cur_doc).exp_data.emax / doc_props(cur_doc).exp_data.delta_e, 0)

first_time = True
redim_time = True
num_cracks = -1
With doc_props(cur_doc)
    With .phisical
        doc_props(cur_doc).results.gs = .es / (2 * (1 + .substrate_pc))
        doc_props(cur_doc).results.lambda = 1 / Sqr((2 * doc_props(cur_doc).results.gs * (.ef * .tf + .es * .ts)) / (.ts ^ 2 * .es * .ef * .tf))
        num_cracks = -1

    End With
    delta = .exp_data.delta_e '* .phisical.lenght * 1000
    
    For ii = 0 To Round(.exp_data.emax / .exp_data.delta_e, 0)
        frm_perform_calculations.pbar.Value = ii
        
        Call delay(0.01)
        With .stress_c_curve
            sigma = .x5 * (ii * delta) ^ 5 + .x4 * (ii * delta) ^ 4 + .x3 * (ii * delta) ^ 3 + .x2 * (ii * delta) ^ 2 + .x * (ii * delta) + .c
            For k = 0 To doc_props(cur_doc).statistic.elements
                If doc_props(cur_doc).elements(k).cracked = False Then
                    doc_props(cur_doc).elements(k).sigma_exp = sigma
                End If
            Next k
        End With
        
        Call new_stress(cur_doc, ii * .exp_data.delta_e)
        last_cracks = doc_props(cur_doc).results.cracks
        Call check_crack(cur_doc, ii * .exp_data.delta_e)
        num_cracks = doc_props(cur_doc).results.cracks
        If last_cracks <> num_cracks Then
                Call new_elements(cur_doc)
        End If
        Call store_data(ii, cur_doc)
        If ii * .exp_data.delta_e * 100 - .exp_data.delta_e * 100 < Round(ii * .exp_data.delta_e * 100, 0) And ii * .exp_data.delta_e * 100 + .exp_data.delta_e * 100 > Round(ii * .exp_data.delta_e * 100, 0) Then
            ' only integer strains: 1,2,3...
            num_cracks = doc_props(cur_doc).results.cracks
            Call new_stress_l_med(cur_doc, Round(ii * .exp_data.delta_e * 100, 0))
        End If
    Next ii
    'Call new_stress(cur_doc, ii * .exp_data.delta_e)
End With
Call redesign_data(cur_doc)
Call delay(0.5)
End Sub
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Sub for re-designing the graphics data
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Sub redesign_data(doc As Integer)

Dim num_flags As Integer
Dim j As Integer
Dim i As Integer
Dim new_graph() As document_results_graphics

With doc_props(doc).results
    For i = 0 To UBound(.graphics) - 1
        For j = i + 1 To UBound(.graphics)
            If .graphics(i).l_med = .graphics(j).l_med Then
                .graphics(j).flag = False
            End If
        Next j
    Next i
    num_flags = 0
    For i = 0 To UBound(.graphics)
        If .graphics(i).flag = True Then
            num_flags = num_flags + 1
        End If
    Next i
    If num_flags < 2 Then
        Exit Sub
    End If
    ReDim new_graph(1 To num_flags)
    j = 1
    For i = 0 To UBound(.graphics)
        If .graphics(i).flag = True Then
            With .graphics(i)
                new_graph(j).crk_density = .crk_density
                new_graph(j).flag = True
                new_graph(j).l_med = .l_med
                new_graph(j).max_elements = doc_props(doc).results.graphics(1).max_elements
                new_graph(j).strain = .strain
                new_graph(j).strain_l_med = .strain_l_med
                ReDim new_graph(j).hl(UBound(.hl))
                new_graph(j).hl = .hl
                
                ReDim new_graph(j).hsf(UBound(.hsf))
                new_graph(j).hsf = .hsf
                
                ReDim new_graph(j).l(UBound(.l))
                new_graph(j).l = .l
                
                ReDim new_graph(j).sf(UBound(.sf))
                new_graph(j).sf = .sf
                
            End With
            j = j + 1
        End If
    Next i
    ReDim .graphics(1 To num_flags)
    For i = 1 To num_flags
        With .graphics(i)
            ReDim .hl(UBound(new_graph(i).hl))
            ReDim .hsf(UBound(new_graph(i).hsf))
            ReDim .l(UBound(new_graph(i).l))
            ReDim .sf(UBound(new_graph(i).sf))
        End With
    Next i
    .graphics = new_graph
End With
End Sub
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Sub for storing data
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Sub store_data(l As Integer, doc As Integer)
Dim k As Integer
With doc_props(doc)
    With .results.live_data(l)
        For k = 1 To doc_props(doc).statistic.elements
            .sigma(k) = doc_props(doc).elements(k).sigma
            .cracks(k) = doc_props(doc).elements(k).cracked
        Next k
    End With
End With
End Sub
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'Function for loading and preparing variables for cracking analysis
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Public Sub Calcular()
Dim cur_doc As Integer
Dim tmp2, tmp, contador As Integer
Dim maxx2 As Double
Dim num_segments As Integer
Dim i As Integer
Dim num_cracks As Integer

num_segments = 30
cur_doc = current_form
With doc_props(cur_doc)
    ReDim .results.live_data(Round(.exp_data.emax / .exp_data.delta_e, 0))
    For i = 0 To Round(.exp_data.emax / .exp_data.delta_e, 0)
        With .results.live_data(i)
            ReDim .cracks(1 To doc_props(cur_doc).statistic.elements)
            ReDim .sigma(1 To doc_props(cur_doc).statistic.elements)
        End With
    Next i
    ' dimensioning the graphics data array
    ReDim .results.graphics(Round(.exp_data.emax * 100, 0))
    For i = 0 To Round(.exp_data.emax * 100, 0)
        With .results.graphics(i)
            ReDim .l(1 To doc_props(cur_doc).statistic.elements)
            ReDim .sf(1 To doc_props(cur_doc).statistic.elements)
            ReDim .hl(1 To 100)
            ReDim .hsf(1 To 100)
            .flag = False
        End With
    Next i
    .results.cracks = 0
    d = .phisical.lenght / .statistic.elements
    If .elements_generated = False Then
        Call stress_load
    Else
        For i = 1 To .statistic.elements
           .elements(i).init = 1
           .elements(i).final = .statistic.elements
           .elements(i).blocks = .statistic.elements
           .elements(i).coord = (i - 1) * d
           .elements(i).cracked = False
           .elements(i).counter = 1
           .elements(i).sigma = .statistic.rs
        Next i
    End If
      Load frm_perform_calculations
      frm_perform_calculations.Show
      frm_perform_calculations.SetFocus
      Call DisableX(frm_perform_calculations)
      Call delay(0.02)
      
      Call engine
      
      num_cracks = 0
      For i = 1 To .statistic.elements
        If .elements(i).cracked Then
            num_cracks = num_cracks + 1
        End If
      Next i
      .results.cracks = num_cracks
      
End With
Unload frm_perform_calculations
FState(cur_doc).calculated = True
Call refresh_richtext
End Sub


