Attribute VB_Name = "Gobal_vars"
Public fMainForm As frmMain

'extension used in the save/open file
Public Const filename_extension = "fsr"

'Specify the files that will appear in the dialog open/save box'
Public Const dialogs_filter = "All Files (*.*)|*.*|FissurMix Files" & "(*." & filename_extension & ")|*." & filename_extension


'build date of the program
Public Const build_date = "2-09-2005"

Public graph_type As String

Type document_exp_data
    x5 As Double
    x4 As Double
    x3 As Double
    x2 As Double
    x As Double
    c As Double
    scales As String
    delta_e As Double
    emax As Double
End Type

Type document_phisical_props
    width_ As Double
    lenght As Double
    substrate_pc As Double
    tf As Double
    ts As Double
    ef As Double
    es As Double
    efr As Double
    esr As Double
    area_s As Double
    area_f As Double
    area_total As Double
End Type

Type document_stat_props
    m As Double
    rs As Double
    s0 As Double
    sl As Double
    elements As Double
End Type

Type document_emodulus_curve
    x4 As Double
    x3 As Double
    x2 As Double
    x As Double
    c As Double
End Type

Type document_stress_curve
    x5 As Double
    x4 As Double
    x3 As Double
    x2 As Double
    x As Double
    c As Double
End Type

Type document_elements
    sigma As Double
    init As Double
    final As Double
    blocks As Double
    coord As Double
    cracked As Boolean
    counter As Integer
    sigmaf As Double
    sigma_exp As Double
End Type

Type document_results_live_data
    sigma() As Double
    cracks() As Boolean
End Type

Type document_results_graphics
    l() As Double ' graphic L_med versus Sigma final - nao homogenizado
    sf() As Double ' graphic L_med versus Sigma final - nao homogenizado
    strain_l_med As Double ' graphic strain versus stress-L_med
    flag As Boolean ' returns true when cracking occurs else (linear elestic behavior) false
    strain As Integer ' extensao
    hl() As Double ' graphic L_med versus Sigma final - homogenizado ao maior Lmed
    hsf() As Double ' graphic L_med versus Sigma final - homogenizado ao maior Lmed
    l_med As Double ' graphic strain versus L_med
    crk_density As Double ' graphic strain versus crack density
    max_elements As Integer ' max de elementos no grafico homogenizado
End Type

Type document_results
    lowest_rnd As Double
    highest_rnd As Double
    sf As Double ' fracture stress
    cracks As Integer
    lambda As Double
    gs As Double
    crack_strain As Double
    live_data() As document_results_live_data
    live_data_pos As Integer
    graphics() As document_results_graphics
End Type

Type document_proprieties
    exp_data As document_exp_data
    statistic As document_stat_props
    phisical As document_phisical_props
    modulus_c_curve As document_emodulus_curve
    stress_c_curve As document_stress_curve
    elements() As document_elements
    elements_generated As Boolean
    results As document_results
End Type

Public doc_props() As document_proprieties

Public frm_exp_data() As New frm_graph_exp_data

