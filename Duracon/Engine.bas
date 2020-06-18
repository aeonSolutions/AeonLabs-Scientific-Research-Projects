Attribute VB_Name = "Engine"
Option Explicit 'Transformas all declared variables into global variables

Type one_form ' holds the current input data entered in the frm_ca_board3
    Distype(5) As Byte
    values As Boolean
    ready As Boolean ' verification of the full data entry for calc
End Type

Type two_form ' holds the current input data entered in the frm_ca_board1
    project_name As String
    Description As String
    project_date As String
    testage_val As String
    testtemp_val As String
    Timeseries_1 As Byte
    cdc As Single
    ready As Boolean ' verification of the full data entry for calc
    values As Boolean
    
End Type

Type simple_form ' holds the current input data entered in the frm_ead_board
    project_name As String
    Description As String
    project_date As String
    Distype(5) As Byte
    ready As Boolean ' verification of the full data entry for calc
    values As Boolean
    
End Type

Type tree_form ' holds the current input data entered in the frm_ca_board2
    Age As String
    Prfsize As Byte
    ready As Boolean ' verification of the full data entry for calc
    values As Boolean
End Type

Type doc_vars
    '*** auxiliares variables
    g As Integer
    '*** permite verificar se um ficheiro foi salvado
    s As Integer
    '*** permite verificar se houve introdução de dados
    '[1-c/introdução  0-s/introdução]
    
    '*** variable description
    predage As Long          '*** Age prediction variable
    prfpred As Integer       '***
    prfsz As Integer         '***
    idifcoef As Integer      '***
    iseedv As Integer        '***
    seed As Long             '***
    nsimul As Long           '***
    tseriev As Integer       '***
    nprojt As String         '*** Project name
    descrip As String        '*** Project description
    datepjt As String          '*** Date of project
    prfvdp(12) As Single     '*** Array for profile depths
    tt As Integer
    kk As Integer
    
    prfvcs(12) As Single     '*** Array for chloride concentrations
    prmdistn(8) As Integer   '*** Array for parameter distribution types 0-2,7
    prmvone(8) As Single     '*** Array for parameter 1 variables
    prmvtwo(8) As Single     '*** Array for parameter 2 variables
    filename As String
    Fileplace As String
    
    sformat As String
    strVariable As String
    filename2 As String
    
    graph_age() As Single     '*** Array for printing age on graph
    graph_pf() As Single       '*** Array for printing pf on graph
    graph_error() As Single    '*** Array for printing error on graph
    graph_beta() As Single     '*** Array for printing beta on graph
    graph_pfbeta() As Single   '*** Array for printing pf of beta on graph
    graph_both() As Single   '*** Array for printing pf and beta on graph
    frm_ca_board3_values As one_form
    frm_ca_board1_values As two_form
    frm_ca_board2_values As tree_form
    frm_ead_board_values As simple_form
End Type

Public doc_props() As doc_vars

Private doc As Integer


'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' Sub for saving the input data in the current active frmchild
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Public Sub save_file(ByRef name, ByRef path, cur_doc As Integer)
  Dim n As Integer
  Dim doc As Integer
  Dim filename As String
  Dim i As Integer
  Dim sformat As String
  
  doc = current_form
  With doc_props(doc)
    ChDir path
    FState(doc).path = path
    filename = name
    i = Len(filename)
    Mid(filename, (i - 2), 3) = "dat"
    Open filename For Output As #1
    
    Print #1, "# D U R A C O N  V" & App.Major & "." & App.Minor & " - (c) 2004 "
    Print #1, " ";
    Print #1, " "
    Print #1, "# Title of the Project "
    Print #1, .nprojt; " - "; .descrip; " ("; .datepjt; ")"; " ;"
    Print #1, " "
    Print #1, " "
    Print #1, "# Model parameters "
    Print #1, "    "; .iseedv; "; # Inital seed (1-randomly generated, 0-user defined) "
    Print #1, "    "; .idifcoef; "; # Diffusion coefficient (0-Dchloride profiling, 1-Dmigration tests, 2-Profile given)"
    Print #1, "    "; .tseriev; "; # Time series (n>0, integer) "
     
    If .idifcoef = 2 Then
        Print #1, "    "; .prfsz; "; # Profile data (5<n<12, interger; 0-no profile) "
        Print #1, "    "; .prfpred; "; # Profile prediction time (0-no, 1-yes) "
    Else
        Print #1, "     0 ; # Profile data (5<n<12, interger; 0-no profile) "
        Print #1, "     0 ; # Profile prediction time (0-no, 1-yes) "
    End If '*** (If idifcoef = 2 Then)***
       
    Print #1, " "
    Print #1, " "
    Print #1, "# Distribution data table "
    Print #1, "#   Distribution types: "
    Print #1, "#     0 - normal "
    Print #1, "#     1 - lognormal "
    Print #1, "#     2 - beta "
    Print #1, "#     7 - deterministic "
    Print #1, "# "
    Print #1, "#         xc       Dcoef         Ccr          Cs           n           t          to           T "
    Print #1, "#       (mm) (1e-12m2/s) (%wt.conc.) (%wt.conc.)         (-)     (years)      (days)  (ºCelcius) "
    Print #1, "# "
    Print #1, "#1_3_5_7_9_1#2_4_6_8_0_2#1_3_5_7_9_1#2_4_6_8_0_2#1_3_5_7_9_1#2_4_6_8_0_2#1_3_5_7_9_1#1_3_5_7_9_1 "
    Print #1, "          "; .prmdistn(0); "         "; .prmdistn(1); "         "; .prmdistn(2); "         "; .prmdistn(3); "         "; .prmdistn(4); "         "; .prmdistn(5); "         "; .prmdistn(6); "         "; .prmdistn(7); ";"
    sformat = "0000.000"
    Print #1, "    "; Replace(Format(CDbl(.prmvone(0)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvone(1)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvone(2)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvone(3)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvone(4)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvone(5)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvone(6)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvone(7)), sformat), ",", "."); " ;"
    Print #1, "    "; Replace(Format(CDbl(.prmvtwo(0)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvtwo(1)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvtwo(2)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvtwo(3)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvtwo(4)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvtwo(5)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvtwo(6)), sformat), ",", "."); "    "; Replace(Format(CDbl(.prmvtwo(7)), sformat), ",", "."); " ;"
       
    Print #1, " "
    Print #1, " "
    Print #1, "# Profile data table "
    Print #1, "# xc (cm) / Cxc (%wt.conc.) "
    Print #1, "# "
    Print #1, "#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7 "
    If .idifcoef = 2 Then
        sformat = "00.00"
        Print #1, "   "; Replace(Format(CDbl(.prfvdp(0)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(1)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(2)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(3)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(4)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(5)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(6)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(7)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(8)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(9)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(10)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvdp(11)), sformat), ",", "."); " ;"
        sformat = "0.000"
        Print #1, "   "; Replace(Format(CDbl(.prfvcs(0)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(1)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(2)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(3)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(4)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(5)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(6)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(7)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(8)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(9)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(10)), sformat), ",", "."); "   "; Replace(Format(CDbl(.prfvcs(11)), sformat), ",", "."); " ;"
    Else
        Print #1, "   00.00   00.00   00.00   00.00   00.00   00.00   00.00   00.00   00.00   00.00   00.00   00.00 ;"
        Print #1, "   0.000   0.000   0.000   0.000   0.000   0.000   0.000   0.000   0.000   0.000   0.000   0.000 ;"
    End If
        
    Print #1, " "
    Print #1, " "
    Print #1, "# Nº Simulations "
    Print #1, "#          n (integer) "
    Print #1, "#1_3_5_7_9_1 "
    Print #1, "   "; .nsimul; "; "
       
    Print #1, " "
    Print #1, " "
    Print #1, "### Initial Seed "
    Print #1, "#          i (integer) "
    Print #1, "#1_3_5_7_9_1 "
    If .iseedv = 0 Then
        Print #1, "   "; .seed; "; "
    Else
        Print #1, "               0 ; "
    End If '*** (If iseedv = 0 Then)***
        
    Print #1, " "
    Print #1, " "
    Print #1, "### Profile prediction time (days) "
    Print #1, "#          i (integer) "
    Print #1, "#1_3_5_7_9_1 "
       
    If .idifcoef = 2 Then
        If .prfpred = 1 Then
            Print #1, "      "; .predage; "; "
        End If
    Else
        Print #1, "           0 ; "
    End If '*** (If idifcoef = 2 Then)***
    Print #1, " "
    Print #1, "# "; .Fileplace
    Print #1, " "
    Print #1, "END_OF_FILE ;"
    
    Close #1 '--->Open name For Output As #1
     
    'abre e guarda um ficheiro com o mesmo nome e extensão .drc
    filename = name
    i = Len(filename)
    Mid(filename, (i - 2), 3) = "drc"
    Open filename For Output As #2
     
    Print #2, .nprojt
    Print #2, .descrip
    Print #2, .datepjt
    Print #2, .iseedv
    Print #2, .idifcoef
    Print #2, .tseriev
    If .idifcoef = 2 Then
        Print #2, .prfsz
        Print #2, .prfpred
    Else
        Print #2, " 0"
        Print #2, " 0"
    End If '*** (If idifcoef = 2 Then)***
    
    Print #2, Replace(.prmdistn(0), ",", "."); "  "; Replace(.prmdistn(1), ",", "."); "  "; Replace(.prmdistn(2), ",", "."); "  "; Replace(.prmdistn(3), ",", "."); "  "; Replace(.prmdistn(4), ",", "."); "  "; Replace(.prmdistn(5), ",", "."); "  "; Replace(.prmdistn(6), ",", "."); "  "; Replace(.prmdistn(7), ",", ".")
    Print #2, Replace(.prmvone(0), ",", "."); "  "; Replace(.prmvone(1), ",", "."); "  "; Replace(.prmvone(2), ",", "."); "  "; Replace(.prmvone(3), ",", "."); "  "; Replace(.prmvone(4), ",", "."); "  "; Replace(.prmvone(5), ",", "."); "  "; Replace(.prmvone(6), ",", "."); "  "; Replace(.prmvone(7), ",", ".")
    Print #2, Replace(.prmvtwo(0), ",", "."); "  "; Replace(.prmvtwo(1), ",", "."); "  "; Replace(.prmvtwo(2), ",", "."); "  "; Replace(.prmvtwo(3), ",", "."); "  "; Replace(.prmvtwo(4), ",", "."); "  "; Replace(.prmvtwo(5), ",", "."); "  "; Replace(.prmvtwo(6), ",", "."); "  "; Replace(.prmvtwo(7), ",", ".")
    Print #2, Replace(.prfvdp(0), ",", "."); "  "; Replace(.prfvdp(1), ",", "."); "  "; Replace(.prfvdp(2), ",", "."); "  "; Replace(.prfvdp(3), ",", "."); "  "; Replace(.prfvdp(4), ",", "."); "  "; Replace(.prfvdp(5), ",", "."); "  "; Replace(.prfvdp(6), ",", "."); "  "; Replace(.prfvdp(7), ",", "."); "  "; Replace(.prfvdp(8), ",", "."); "  "; Replace(.prfvdp(9), ",", "."); "  "; Replace(.prfvdp(10), ",", "."); "  "; Replace(.prfvdp(11), ",", ".")
    Print #2, Replace(.prfvcs(0), ",", "."); "  "; Replace(.prfvcs(1), ",", "."); "  "; Replace(.prfvcs(2), ",", "."); "  "; Replace(.prfvcs(3), ",", "."); "  "; Replace(.prfvcs(4), ",", "."); "  "; Replace(.prfvcs(5), ",", "."); "  "; Replace(.prfvcs(6), ",", "."); "  "; Replace(.prfvcs(7), ",", "."); "  "; Replace(.prfvcs(8), ",", "."); "  "; Replace(.prfvcs(9), ",", "."); "  "; Replace(.prfvcs(10), ",", "."); "  "; Replace(.prfvcs(11), ",", ".")
        
    If .iseedv = 0 Then
        Print #2, .seed
    Else
        Print #2, "0"
    End If '*** (If iseedv = 0 Then)***
       
    Print #2, .nsimul
     
    If .idifcoef = 2 Then
        If .prfpred = 1 Then
            Print #2, .predage
        End If
    Else
        Print #2, " 0"
    End If '*** (If idifcoef = 2 Then)***
    
    Close #2 '--->Open FileName For Output As #2
End With
FState(doc).saved = True
document(doc).Caption = name
FState(doc).name = name
End Sub

Private Sub add_item(doc As Integer, rowo As Integer, colo As Integer, texto As String, merge As Boolean, bold As Boolean)

Dim i As Integer
Dim lenght As Integer
Dim tmp() As String

If Len(texto) > 3 Then
    lenght = Int(Len(texto) / 15)
Else
    lenght = 0
End If
tmp() = Split(texto, ",")
If UBound(tmp) > 0 Then
    texto = tmp(0) & "." & tmp(1)
End If
If texto = "0" Then
    texto = "0.0"
End If

document(doc).lista.ColWidth(colo) = document(doc).TextWidth("########") * 2

For i = colo To colo + lenght
    With document(doc)
        .lista.CellAlignment = flexAlignLeftCenter
        .lista.Row = rowo
        .lista.Col = i
        .lista.Text = texto
        If merge Then
            .lista.MergeRow(rowo) = True
        End If
        If bold Then
            .lista.CellFontBold = True
        End If
    End With
Next i
End Sub

Public Sub refresh_lista(doc As Integer)
        
Dim i As Integer
Dim tmp As String
Dim tmp2(4) As String
Dim tmp3 As Integer

document(doc).lista.Enabled = False
document(doc).lista.Visible = False
document(doc).lista.Clear
With doc_props(doc)
        Call add_item(doc, 1, 1, "Project Name:", True, True)
        If .nprojt <> "" Then
            Call add_item(doc, 2, 2, .nprojt, True, False)
        Else
            Call add_item(doc, 2, 2, "- None -", True, False)
        End If
        Call add_item(doc, 3, 1, "Description:", True, True)
        If .descrip <> "" Then
            Call add_item(doc, 4, 2, .descrip, True, False)
        Else
            Call add_item(doc, 4, 2, "- None -", True, False)
        End If
        Call add_item(doc, 5, 1, "Date:", True, True)
        If .descrip <> "" Then
            Call add_item(doc, 6, 2, .datepjt, True, False)
        Else
            Call add_item(doc, 6, 2, "- None -", True, False)
        End If
        tmp3 = 8
        If .frm_ead_board_values.values = True Then
            Call add_item(doc, 8, 1, "Model Parameters:", True, True)
            Call add_item(doc, 9, 2, "Chloride Diffusion Coefficient:", True, False)
            Call add_item(doc, 10, 3, "Design Value", True, False)
            Call add_item(doc, 12, 2, "Design Life of Structure :  50 years", True, False)
            Call add_item(doc, 14, 2, "Age of Structure during assessment :  28 days", True, False)
            tmp3 = 18
        End If
        If .frm_ca_board1_values.values = True Then
            Call add_item(doc, 8, 1, "Model Parameters:", True, True)
            Call add_item(doc, 9, 2, "Chloride Diffusion Coefficient:", True, False)
            If doc_props(doc).frm_ca_board1_values.cdc = 4 Then
                Call add_item(doc, 10, 3, "Design Value", True, False)
            ElseIf doc_props(doc).frm_ca_board1_values.cdc = 1 Then
                Call add_item(doc, 10, 3, "Obtained from testing - NT Build 492", True, False)
            ElseIf doc_props(doc).frm_ca_board1_values.cdc = 2 Then
                Call add_item(doc, 10, 3, "Obtained from testing - NT Build 443", True, False)
            ElseIf doc_props(doc).frm_ca_board1_values.cdc = 3 Then
                Call add_item(doc, 10, 3, "Obtained from chloride profile", True, False)
            End If
            If .frm_ca_board1_values.Timeseries_1 = 0 Then
                tmp = "50"
            ElseIf .frm_ca_board1_values.Timeseries_1 = 1 Then
                tmp = "75"
            ElseIf .frm_ca_board1_values.Timeseries_1 = 2 Then
                tmp = "100"
            ElseIf .frm_ca_board1_values.Timeseries_1 = 3 Then
                tmp = "125"
            ElseIf .frm_ca_board1_values.Timeseries_1 = 4 Then
                tmp = "150"
            End If
            Call add_item(doc, 12, 2, "Design Life of Structure :  " & tmp & " years", True, False)
            If .frm_ca_board1_values.testage_val = "N/A" Then
                Call add_item(doc, 14, 2, "Age of Structure during assessment :  28 days", True, False)
            Else
                Call add_item(doc, 14, 2, "Age of Structure during assessment :  " & .frm_ca_board1_values.testage_val & " days", True, False)
            End If
            If doc_props(doc).frm_ca_board1_values.cdc = 3 And .frm_ca_board2_values.values Then
                Call add_item(doc, 18, 1, "Chloride profile Information", True, True)
                Call add_item(doc, 19, 2, "Profile predition:", True, False)
                If .frm_ca_board2_values.Age = "N/A" Then
                    Call add_item(doc, 19, 3, "No", True, False)
                Else
                    Call add_item(doc, 19, 3, "Yes ( " & .frm_ca_board2_values.Age & " days )", True, False)
                End If
                Call add_item(doc, 21, 2, "Profile Values:", True, False)
                Call add_item(doc, 23, 2, "Depth (cm)", True, True)
                Call add_item(doc, 24, 2, "Cl Concentration(%)", True, True)
                Call add_item(doc, 22, 2, "     nº: ", True, True)
                For i = 0 To 4 + .frm_ca_board2_values.Prfsize
                    Call add_item(doc, 23, 4 + i, Str(.prfvdp(i)), False, False)
                    Call add_item(doc, 22, 4 + i, Str(i + 1), False, True)
                    Call add_item(doc, 24, 4 + i, Str(.prfvcs(i)), False, False)
                Next i
                tmp3 = 26
            Else
                tmp3 = 18
            End If
        End If
        If .frm_ca_board3_values.values Or .frm_ead_board_values.values Then
            tmp2(0) = "Deterministic"
            tmp2(1) = "Normal"
            tmp2(2) = "Lognormal"
            tmp2(3) = "Beta"
            Call add_item(doc, tmp3, 1, "Distribuition Data", True, True)
            
            'concrete cover
            Call add_item(doc, tmp3 + 2, 4, tmp2(.frm_ca_board3_values.Distype(0)), False, False)
            Call add_item(doc, tmp3 + 2, 5, " " & Str(.prmvone(0)), False, False)
            Call add_item(doc, tmp3 + 2, 6, Str(.prmvtwo(0)), False, False)
            
            'diffusion coefficient
            Call add_item(doc, tmp3 + 3, 4, tmp2(.frm_ca_board3_values.Distype(1)), False, False)
            Call add_item(doc, tmp3 + 3, 5, " " & Str(.prmvone(1)), False, False)
            Call add_item(doc, tmp3 + 3, 6, Str(.prmvtwo(1)), False, False)
            
            'critial cl concentration
            Call add_item(doc, tmp3 + 4, 4, tmp2(.frm_ca_board3_values.Distype(2)), False, False)
            Call add_item(doc, tmp3 + 4, 5, " " & Str(.prmvone(2)), False, False)
            Call add_item(doc, tmp3 + 4, 6, Str(.prmvtwo(2)), False, False)
            
            'surface cl concentration
            Call add_item(doc, tmp3 + 5, 4, tmp2(.frm_ca_board3_values.Distype(3)), False, False)
            Call add_item(doc, tmp3 + 5, 5, " " & Str(.prmvone(3)), False, False)
            Call add_item(doc, tmp3 + 5, 6, Str(.prmvtwo(3)), False, False)
            
            ' age effect on diffusion
            Call add_item(doc, tmp3 + 6, 4, tmp2(.frm_ca_board3_values.Distype(4)), False, False)
            Call add_item(doc, tmp3 + 6, 5, " " & Str(.prmvone(4)), False, False)
            Call add_item(doc, tmp3 + 6, 6, Str(.prmvtwo(4)), False, False)
        
            Call add_item(doc, tmp3 + 1, 4, "Type", True, True)
            Call add_item(doc, tmp3 + 1, 5, "Parameter 1", True, True)
            Call add_item(doc, tmp3 + 1, 6, "Parameter 2", True, True)
            
            Call add_item(doc, tmp3 + 2, 2, "Concrete cover", True, True)
            Call add_item(doc, tmp3 + 3, 2, "Diffusion Coefficient", True, True)
            Call add_item(doc, tmp3 + 4, 2, "Critical Cl concentration", True, True)
            Call add_item(doc, tmp3 + 5, 2, "Surface Cl concentration", True, True)
            Call add_item(doc, tmp3 + 6, 2, "Age effect diffusion", True, True)
        End If
End With
document(doc).lista.Visible = True
document(doc).lista.Enabled = True
document(doc).lista.Refresh
End Sub

Public Sub unload_document()
Dim doc As Integer
Dim tmp As VbMsgBoxResult
Dim j As Integer
Dim name As String
Dim path As String
Dim i As Integer

doc = current_form()
If FState(doc).values Then
  If Not FState(doc).saved Then
    tmp = MsgBox("Save the Document ?", vbYesNoCancel + vbCritical, "Duracon")
    If tmp = vbCancel Then
      Exit Sub
    End If
    If tmp = vbYes Then
           ' Set CancelError is True
           document(doc).Dialogs.CancelError = True
           On Error Resume Next
           ' Set flags
           document(doc).Dialogs.Flags = cdlOFNHideReadOnly
           ' Set filters
           document(doc).Dialogs.Filter = "All Files (*.*)|*.*|Duracon Files" & "(*.drc)|*.drc"
           ' Specify default filter
           document(doc).Dialogs.FilterIndex = 2
           ' set the working directory the application dir
           document(doc).Dialogs.InitDir = App.path
           ' Display the save dialog box
           document(doc).Dialogs.ShowSave
           If Err.Number <> 0 Then
             Exit Sub
           End If
           ' get the name file and the path
           name = GetFile(document(doc).Dialogs.filename)
           path = GetPath(document(doc).Dialogs.filename)
         Call save_file(name, path, doc)
    End If
  End If
End If
FState(doc).saved = False
FState(doc).values = False
With doc_props(doc)
    .frm_ca_board1_values.ready = False
    .frm_ca_board2_values.ready = False
    .frm_ca_board3_values.ready = False
    .frm_ead_board_values.ready = False
    .frm_ca_board1_values.values = False
    .frm_ca_board2_values.values = False
    .frm_ca_board3_values.values = False
    .frm_ead_board_values.values = False
    .nprojt = ""
    .datepjt = ""
    .descrip = ""
    .predage = 0
    .prfpred = 0
    .prfsz = 0
    .idifcoef = 0
    .iseedv = 0
    .seed = 0
    .nsimul = 0
    .tseriev = 0
For i = 0 To UBound(.prfvdp)
    .prfvdp(i) = 0
    .prfvcs(i) = 0
Next i
For i = 0 To UBound(.prfvdp)
        .prfvdp(i) = 0
    Next i
    .tt = 0
    .kk = 0
    For i = 0 To UBound(.prfvcs)
        .prfvcs(i) = 0
    Next i
    For i = 0 To UBound(.prmdistn)
        .prmdistn(i) = 0
    Next i
    For i = 0 To UBound(.prmvone)
        .prmvone(i) = 0
        .prmvtwo(i) = 0
    Next i
    Call refresh_lista(doc)
End With
With document(doc)
    .pf_chart.Data = 0
    .reliability_chart.Data = 0
End With
FState(doc).deleted = True
Unload document(doc)
Unload pf_graph(doc)
Unload ri_graph(doc)
FState(doc).deleted = True
FState(doc).Dirty = False
'If doc > 1 Then
'    FState(doc - 1).Dirty = True
'End If

End Sub

