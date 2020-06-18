Attribute VB_Name = "General"
Option Explicit

Sub reset_props()
Dim doc As Integer

doc = current_form
With doc_props(doc)
    With .exp_data
        .scales = "-1"
        .x5 = -1
        .x4 = -1
        .x3 = -1
        .x2 = -1
        .x = -1
        .c = -1
        .emax = -1
        .delta_e = -1
    End With
    With .phisical
        .lenght = -1
        .substrate_pc = -1
        .tf = -1
        .ts = -1
        .ef = -1
    End With
    With .statistic
        .m = -1
        .rs = -1
    End With
End With
FState(doc).values = False
End Sub

Sub refresh_richtext()
Dim tmp, tmp2 As String
Dim doc As Integer
Dim i As Integer
Dim average_beta As Double


doc = current_form
If doc = -1 Then
    Exit Sub
End If
document(doc).RichTextBox.LoadFile App.path & "\report.rtf"
tmp = document(doc).RichTextBox.TextRTF
If FState(doc).name = "" Then
    tmp = Replace(tmp, "[CODED]", "Unsaved document", 1, 1, vbTextCompare)
Else
    tmp = Replace(tmp, "[CODED]", FState(doc).name, 1, 1, vbTextCompare)
End If
If FState(doc).values Then
    'construct the polynomial load curve
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.x5, 5)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.x4, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.x3, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.x2, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.x, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.c, 3)) & " [" & doc_props(doc).exp_data.scales & "]", 1, 1, vbTextCompare)
    'construct the stress curve
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).stress_c_curve.x5 / 1000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).stress_c_curve.x4 / 1000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).stress_c_curve.x3 / 1000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).stress_c_curve.x2 / 1000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).stress_c_curve.x / 1000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).stress_c_curve.c / 1000000#, 3)) & " [MPa]", 1, 1, vbTextCompare)
    'construct the elastic modulus curve
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).modulus_c_curve.x4 / 1000000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).modulus_c_curve.x3 / 1000000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).modulus_c_curve.x2 / 1000000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).modulus_c_curve.x / 1000000000#, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).modulus_c_curve.c / 1000000000#, 3)) & " [GPa]", 1, 1, vbTextCompare)
    
Else
    ' replace polynomial load curve
    tmp = Replace(tmp, "[CODED] x\super 5\nosupersub  + [CODED] x\super 4\nosupersub  + [CODED] x\super 3\nosupersub  + [CODED] x\super 2\nosupersub  + [CODED] x + [CODED]", "You need to load experimental data first!", 1, 1, vbTextCompare)
    'replace the stress curve
    tmp = Replace(tmp, "[CODED] x\super 5\nosupersub  + [CODED] x\super 4\nosupersub  + [CODED] x\super 3\nosupersub  + [CODED] x\super 2\nosupersub  + [CODED] x + [CODED]", " - - ", 1, 1, vbTextCompare)
    'replace the elastic modulus curve
    tmp = Replace(tmp, "[CODED] x\super 4\nosupersub  + [CODED] x\super 3\nosupersub  + [CODED] x\super 2\nosupersub  + [CODED] x + [CODED]", " - - ", 1, 1, vbTextCompare)

End If
If FState(doc).values Then
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.emax * 100, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.delta_e * 100, 3)), 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", CStr(Round(doc_props(doc).exp_data.emax / doc_props(doc).exp_data.delta_e, 0)), 1, 1, vbTextCompare)
Else
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
End If

If FState(doc).values Then
    With doc_props(doc).phisical
        tmp2 = CStr(.lenght * 1000) ' lenght
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        
        tmp2 = CStr(.width_ * 1000)    ' width
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        
        tmp2 = CStr(.tf * 1000)  ' Tf
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        
        tmp2 = CStr(.ef / 1000000000#)  ' Ef
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        
        tmp2 = CStr(.efr * 100)  ' film residual strain efr
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
            
        tmp2 = CStr(.ts * 1000)  ' Ts
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        
        tmp2 = CStr(.es / 1000000000#)  ' Es
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        
        tmp2 = CStr(.esr * 100)  ' substrate residual strain efr
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        
        tmp2 = CStr(.substrate_pc)  ' vs
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
    End With
    With doc_props(doc).statistic
        tmp2 = CStr(.m)   ' m
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        tmp2 = CStr(.s0 / 1000000#)  ' s0
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        tmp2 = CStr(.sl / 1000000#) ' sl
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        tmp2 = CStr(.rs / 1000000#) ' rs
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
        tmp2 = CStr(.elements)  ' number of elements
        tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
    End With
Else
    For i = 1 To 14
        tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
    Next i
End If

If FState(doc).calculated Then
    tmp = Replace(tmp, "[CODED]", CStr(doc_props(doc).results.cracks), 1, 1, vbTextCompare)
Else
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
End If

If doc_props(doc).elements_generated Then
    tmp2 = doc_props(doc).results.lowest_rnd / 1000000#
    If tmp2 > 10000 Then
        tmp2 = tmp2 / 1000  ' converted to GPa
        tmp2 = CStr(Round(tmp2, 3)) & " GPa"
    Else
        tmp2 = CStr(Round(tmp2, 3)) & " MPa"
    End If
    tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
Else
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
End If
If doc_props(doc).elements_generated Then
    tmp2 = doc_props(doc).results.highest_rnd / 1000000#
    If tmp2 > 10000 Then
        tmp2 = tmp2 / 1000  ' converted to GPa
        tmp2 = CStr(Round(tmp2, 3)) & " GPa"
    Else
        tmp2 = CStr(Round(tmp2, 3)) & " MPa"
    End If
    tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
Else
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
End If


If FState(doc).calculated Then ' fracture stress and fracture strain
    tmp2 = doc_props(doc).results.sf / 1000000#
    If tmp2 > 10000 Then
        tmp2 = tmp2 / 1000  ' converted to GPa
        tmp2 = CStr(Round(tmp2, 3)) & " GPa"
    Else
        tmp2 = CStr(Round(tmp2, 3)) & " MPa"
    End If
    tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
    
    tmp2 = doc_props(doc).results.crack_strain * 100
    tmp = Replace(tmp, "[CODED]", tmp2, 1, 1, vbTextCompare)
Else
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
    tmp = Replace(tmp, "[CODED]", " - - ", 1, 1, vbTextCompare)
End If

document(doc).RichTextBox.TextRTF = tmp
End Sub

