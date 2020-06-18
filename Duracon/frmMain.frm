VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Duracon"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Duracon"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   480
      Top             =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2619
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "04-04-2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "14:07"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialogs 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_print 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnufilebar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_insert_data 
         Caption         =   "Insert Data"
         Enabled         =   0   'False
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_clear_data 
         Caption         =   "Clear Data"
      End
   End
   Begin VB.Menu mnu_simulation 
      Caption         =   "&Simulation"
      Begin VB.Menu mnu_run 
         Caption         =   "Run analysis"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnu_results 
      Caption         =   "&Graphics"
      Begin VB.Menu mnu_reliability 
         Caption         =   "Reliability Index vs. Time"
      End
      Begin VB.Menu mnu_probability 
         Caption         =   "&Probability of Failure vs. Time"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Report a Bug"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    'FileNew
    Me.mnu_results.Enabled = False
    Me.mnu_run.Enabled = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub


Private Sub mnu_clear_data_Click()
Dim doc As Integer
Dim arraycount As Integer
Dim tmp As Boolean
Dim i As Integer

On Error Resume Next
arraycount = UBound(document)
If Err.Number <> 0 Then
    tmp = False
Else ' documents found - check if they really exist
    doc = current_form
    tmp = False
    If Not FState(doc).deleted Then
        tmp = True
    End If
End If

If Not tmp Then
    MsgBox "You need to have at least one document open!", vbOKOnly + vbCritical, " Duracon "
    Exit Sub
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
End Sub

Private Sub mnu_insert_data_Click()
Dim doc As Integer
Dim arraycount As Integer
Dim i As Integer
Dim tmp As Boolean

On Error Resume Next
arraycount = UBound(document)
If Err.Number <> 0 Then
    tmp = False
Else ' documents found - check if they really exist
    tmp = False
    For i = 1 To arraycount
        If Not FState(i).deleted Then
            tmp = True
            Exit For
        End If
    Next i
End If

If Not tmp Then
    MsgBox "You need to have at least one document open!", vbOKOnly + vbCritical, " Duracon "
    Exit Sub
End If

doc = current_form
With doc_props(doc)
    .idifcoef = 1
    .iseedv = 1
    .prfpred = 0
    .tt = 0
End With
tmp = False
If Not (doc_props(doc).frm_ca_board3_values.ready And doc_props(doc).frm_ca_board1_values.ready) Then
    If doc_props(doc).frm_ca_board1_values.cdc = 3 Then
        If doc_props(doc).frm_ca_board2_values.ready Then
            tmp = False
        Else
            tmp = True
        End If
    Else
        tmp = True
    End If
End If

If tmp Then ' no valid data found
    frm_choice.Show 1
Else
    If doc_props(doc).frm_ca_board1_values.cdc = -1 Then
        frm_ead_board.Show 1
    Else
        frm_ca_board1.Show 1
    End If
End If
tmp = False
If Not (doc_props(doc).frm_ca_board3_values.ready And doc_props(doc).frm_ca_board1_values.ready) Then
    If doc_props(doc).frm_ca_board1_values.cdc = 3 Then
        If doc_props(doc).frm_ca_board2_values.ready Then
            tmp = False
        Else
            tmp = True
        End If
    Else
        tmp = True
    End If
End If

If tmp Then 'no valid data found
    Exit Sub
End If
If FState(doc).saved = False Then ' first you have to save the file
    Exit Sub
End If
Me.mnu_run.Enabled = True
Me.mnu_insert_data.Caption = "Edit Data"

End Sub

Private Sub mnu_print_Click()
Dim doc As Integer
Dim sformat As String

doc = current_form
With doc_props(doc)

   Printer.Print "# D U R A C O N  v1.0 - (c) 2004 "
   Printer.Print " "
   Printer.Print " "
   Printer.Print "# Title of the Project "
   Printer.Print .nprojt; " - "; .descrip; " ("; .datepjt; ")"; " ;"
   Printer.Print " "
   Printer.Print " "
   Printer.Print "# Model parameters "
   Printer.Print "    "; .iseedv; "; # Inital seed (1-randomly generated, 0-user defined) "
   Printer.Print "    "; .idifcoef; "; # Diffusion coefficient (0-Dchloride profiling, 1-Dmigration tests, 2-Profile given)"
   Printer.Print "    "; .tseriev; "; # Time series (n>0, integer) "
   
    If .idifcoef = 2 Then
        Printer.Print "    "; .prfsz; "; # Profile data (5<n<12, interger; 0-no profile) "
        Printer.Print "    "; .prfpred; "; # Profile prediction time (0-no, 1-yes) "
    Else
        Printer.Print "     0 ; # Profile data (5<n<12, interger; 0-no profile) "
        Printer.Print "     0 ; # Profile prediction time (0-no, 1-yes) "
    End If '*** (If idifcoef = 2 Then)***
   
   Printer.Print " "
   Printer.Print " "
   Printer.Print "# Distribution data table "
   Printer.Print "#   Distribution types: "
   Printer.Print "#     0 - normal "
   Printer.Print "#     1 - lognormal "
   Printer.Print "#     2 - beta "
   Printer.Print "#     7 - deterministic "
   Printer.Print "# "
   Printer.Print "#         xc       Dcoef         Ccr          Cs           n           t          to           T "
   Printer.Print "#       (mm) (1e-12m2/s) (%wt.conc.) (%wt.conc.)         (-)     (years)      (days)  (ºCelcius) "
   Printer.Print "# "
   Printer.Print "#1_3_5_7_9_1#2_4_6_8_0_2#1_3_5_7_9_1#2_4_6_8_0_2#1_3_5_7_9_1#2_4_6_8_0_2#1_3_5_7_9_1#1_3_5_7_9_1 "
   Printer.Print "          "; .prmdistn(0); "         "; .prmdistn(1); "         "; .prmdistn(2); "         "; .prmdistn(3); "         "; .prmdistn(4); "         "; .prmdistn(5); "         "; .prmdistn(6); "         "; .prmdistn(7); ";"
   sformat = "0000.000"
   Printer.Print "    "; Format(Val(.prmvone(0)), sformat); "    "; Format(Val(.prmvone(1)), sformat); "    "; Format(Val(.prmvone(2)), sformat); "    "; Format(Val(.prmvone(3)), sformat); "    "; Format(Val(.prmvone(4)), sformat); "    "; Format(Val(.prmvone(5)), sformat); "    "; Format(Val(.prmvone(6)), sformat); "    "; Format(Val(.prmvone(7)), sformat); " ;"
   Printer.Print "    "; Format(Val(.prmvtwo(0)), sformat); "    "; Format(Val(.prmvtwo(1)), sformat); "    "; Format(Val(.prmvtwo(2)), sformat); "    "; Format(Val(.prmvtwo(3)), sformat); "    "; Format(Val(.prmvtwo(4)), sformat); "    "; Format(Val(.prmvtwo(5)), sformat); "    "; Format(Val(.prmvtwo(6)), sformat); "    "; Format(Val(.prmvtwo(7)), sformat); " ;"
   
   Printer.Print " "
   Printer.Print " "
   Printer.Print "# Profile data table "
   Printer.Print "# xc (cm) / Cxc (%wt.conc.) "
   Printer.Print "# "
   Printer.Print "#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7#1_3_5_7 "
   sformat = "00.00"
   Printer.Print "   "; Format(Val(.prfvdp(0)), sformat); "   "; Format(Val(.prfvdp(1)), sformat); "   "; Format(Val(.prfvdp(2)), sformat); "   "; Format(Val(.prfvdp(3)), sformat); "   "; Format(Val(.prfvdp(4)), sformat); "   "; Format(Val(.prfvdp(5)), sformat); "   "; Format(Val(.prfvdp(6)), sformat); "   "; Format(Val(.prfvdp(7)), sformat); "   "; Format(Val(.prfvdp(8)), sformat); "   "; Format(Val(.prfvdp(9)), sformat); "   "; Format(Val(.prfvdp(10)), sformat); "   "; Format(Val(.prfvdp(11)), sformat); ";"
   sformat = "0.000"
   Printer.Print "   "; Format(Val(.prfvcs(0)), sformat); "   "; Format(Val(.prfvcs(1)), sformat); "   "; Format(Val(.prfvcs(2)), sformat); "   "; Format(Val(.prfvcs(3)), sformat); "   "; Format(Val(.prfvcs(4)), sformat); "   "; Format(Val(.prfvcs(5)), sformat); "   "; Format(Val(.prfvcs(6)), sformat); "   "; Format(Val(.prfvcs(7)), sformat); "   "; Format(Val(.prfvcs(8)), sformat); "   "; Format(Val(.prfvcs(9)), sformat); "   "; Format(Val(.prfvcs(10)), sformat); "   "; Format(Val(.prfvcs(11)), sformat); ";"
   
   Printer.Print " "
   Printer.Print " "
   Printer.Print "# Nº Simulations "
   Printer.Print "#          n (integer) "
   Printer.Print "#1_3_5_7_9_1 "
   Printer.Print "   "; .nsimul; "; "
     
   Printer.Print " "
   Printer.Print " "
   Printer.Print "### Initial Seed "
   Printer.Print "#          i (integer) "
   Printer.Print "#1_3_5_7_9_1 "
    If .iseedv = 0 Then
        Printer.Print "   "; .seed; "; "
    Else
        Printer.Print "               0 ; "
    End If '*** (If iseedv = 0 Then)***
   
   Printer.Print " "
   Printer.Print " "
   Printer.Print "### Profile prediction time (days) "
   Printer.Print "#          i (integer) "
   Printer.Print "#1_3_5_7_9_1 "
   
    If .idifcoef = 2 Then
        If .prfpred = 1 Then
            Printer.Print "      "; .predage; "; "
        End If
    Else
        Printer.Print "           0 ; "
    End If '*** (If idifcoef = 2 Then)***
  
   Printer.Print " "
   Printer.Print "END_OF_FILE ;"
End With
End Sub

Private Sub mnu_probability_Click()
Dim doc As Integer
doc = current_form
pf_graph(doc).Show 1
End Sub

Private Sub mnu_reliability_Click()
Dim doc As Integer
doc = current_form
ri_graph(doc).Show 1
End Sub

Private Sub mnu_run_Click()
Dim i As Integer
Dim j As Integer
Dim mm As Integer
Dim vals As Single
Dim aux1 As Single
Dim aux2 As Single
Dim numdt As Single
Dim maxage As Single
Dim mgraph As Single
Dim bgraph As Single
Dim betamax As Single
Dim betamin As Single
Dim doc As Integer
Dim tmp As Boolean
Dim arraycount As Integer
Dim exec_file_name As String
Dim r() As String
Dim file_path As String
Dim filename As String
Dim retval As Long
Dim m As Double
Dim b As Double
Dim time_scale As Integer

On Error Resume Next
arraycount = UBound(document)
If Err.Number <> 0 Then
    tmp = False
Else ' documents found - check if they really exist
    tmp = False
    For i = 1 To arraycount
        If Not FState(i).deleted Then
            tmp = True
            Exit For
        End If
    Next i
End If

If Not tmp Then
    MsgBox "You need to have at least one document open!", vbOKOnly + vbCritical, " Duracon "
    Exit Sub
End If

doc = current_form
tmp = False
If doc_props(doc).frm_ead_board_values.ready Then
 tmp = False
ElseIf Not (doc_props(doc).frm_ca_board3_values.ready And doc_props(doc).frm_ca_board1_values.ready) Then
    If doc_props(doc).frm_ca_board1_values.cdc = 3 Then
        If doc_props(doc).frm_ca_board2_values.ready Then
            tmp = False
        Else
            tmp = True
        End If
    Else
        tmp = True
    End If
End If

If tmp Then
    MsgBox "There are some data entry missing.Please verify input data! ", vbOKOnly + vbCritical, " Duracon "
    Exit Sub
End If

If FState(doc).saved = False Then
    MsgBox "You need to save the file first.", vbOKOnly + vbCritical, " Duracon "
    Exit Sub
End If
Screen.MousePointer = vbHourglass
Load frm_perform_calculations
frm_perform_calculations.Show
frm_perform_calculations.SetFocus
Call DisableX(frm_perform_calculations)
Call delay(0.02)
filename = FState(doc).name
r() = Split(filename, ".")
filename = r(0)

file_path = FState(doc).path
' call the form that handles the data preps and caculations
frm_perform_calculations.txt.Caption = "Performing caculations.... One moment, please!"
' preparing data to be run in the core - system\duracon.sys
'   copy the current document file to system folder and name it simulation.dat
Call Copy_file(file_path & "\" & filename & ".dat", App.path & "\system\simulation.dat")
'   copy and rename duracon.sys to duracon.bat
Call Copy_file(App.path & "\system\duracon.sys", App.path & "\system\runner.exe")
'   run duracon.bat
ChDir (App.path & "\system")
retval = ExecCmd("runner.exe simulation")
'   copy and rename the result files to the document's folder
Call Copy_file(App.path & "\system\simulation_rs.dat", file_path & "\" & filename & "_rs.dat")
Call Copy_file(App.path & "\system\simulation_xls.dat", file_path & "\" & filename & "_xls.dat")
'   delete files duracon.bat, simulation.bat
Kill App.path & "\system\runner.exe"
Kill App.path & "\system\simulation_xls.dat"
Kill App.path & "\system\simulation_rs.dat"
Kill App.path & "\system\simulation.dat"

file_path = FState(doc).path
r() = Split(FState(doc).name, ".")
filename = r(0)

'opening data file to print graphs!
ChDir FState(doc).path
r() = Split(FState(doc).name, ".")
filename = r(0) & "_xls.dat"
On Error Resume Next
Open filename For Input As #1
If Err.Number <> 0 Then ' file not found!?
    MsgBox "Results file not found! ", vbOKOnly + vbCritical, " Duracon "
    frm_perform_calculations.Hide
    Unload frm_perform_calculations
    Screen.MousePointer = vbDefault
    Exit Sub
End If

pf_graph(doc).Caption = "Probability of Failure Curve: " & document(doc).Caption
ri_graph(doc).Caption = "Reliability Index Curve: " & document(doc).Caption


On Error GoTo 0
    
Input #1, vals
maxage = vals
If maxage = 50 Then
    time_scale = 5
ElseIf maxage = 75 Then
    time_scale = 8
ElseIf maxage = 100 Then
    time_scale = 10
ElseIf maxage = 125 Then
    time_scale = 12
ElseIf maxage = 150 Then
    time_scale = 5
End If
numdt = (maxage / 5) + 1
ReDim doc_props(doc).graph_pf(1 To numdt + 1, 1 To 2)
ReDim doc_props(doc).graph_error(1 To numdt, 1 To 2)
ReDim doc_props(doc).graph_beta(1 To numdt + 1, 1 To 2)
ReDim doc_props(doc).graph_pfbeta(1 To numdt, 1 To 2)

doc_props(doc).graph_pf(i + 1, 1) = 0
doc_props(doc).graph_pf(i + 1, 2) = 0

For i = 1 To numdt
    Input #1, vals
    doc_props(doc).graph_pf(i + 1, 1) = vals
    doc_props(doc).graph_error(i, 1) = vals
    doc_props(doc).graph_beta(i + 1, 1) = vals
    doc_props(doc).graph_pfbeta(i, 1) = vals
    Input #1, vals
    doc_props(doc).graph_pf(i + 1, 2) = vals
    Input #1, vals
    doc_props(doc).graph_error(i, 2) = vals
    Input #1, vals
    doc_props(doc).graph_beta(i + 1, 2) = vals
    Input #1, vals
    doc_props(doc).graph_pfbeta(i, 2) = vals
Next i
'y(x)=mx+b
With doc_props(doc)
    m = (.graph_beta(2, 2) - .graph_beta(3, 2)) / (.graph_beta(2, 1) - .graph_beta(3, 1))
    b = .graph_beta(3, 2) - m * .graph_beta(3, 1)
    
    .graph_beta(1, 1) = 0
    .graph_beta(1, 2) = b
End With
Close #1

' PLOTTING GRAPHICS
frm_perform_calculations.txt.Caption = "Plotting graphics.... One moment, please!"
frm_perform_calculations.pbar.Max = 6
frm_perform_calculations.pbar.Value = 0
'************************************
'DRAWING GRAPH Probability of Failure
'************************************
With document(doc)
    With .pf_chart
        .RowCount = numdt
        .ChartData = doc_props(doc).graph_pf()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = "Probability of failure (%)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = False
            .ValueScale.Maximum = 100
            .ValueScale.Minimum = 0
        End With

        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = "Time (years)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = False
            .AxisGrid.MajorPen.Style = VtPenStyleNull
            .AxisGrid.MinorPen.Style = VtPenStyleNull
            
            .ValueScale.MajorDivision = time_scale
            .ValueScale.MinorDivision = 0
            
            .ValueScale.Maximum = maxage
            .ValueScale.Minimum = 5
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        frm_perform_calculations.pbar.Value = 1
        Call delay(1#)
        .Refresh
    End With
With pf_graph(doc)
    With .pf_chart
        .RowCount = numdt
        .ChartData = doc_props(doc).graph_pf()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = "Probability of failure (%)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = False
            .ValueScale.Maximum = 100
            .ValueScale.Minimum = 0
        End With

        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = "Time (years)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = False
            .AxisGrid.MajorPen.Style = VtPenStyleNull
            .AxisGrid.MinorPen.Style = VtPenStyleNull
            
            .ValueScale.MajorDivision = time_scale
            .ValueScale.MinorDivision = 0
            
            .ValueScale.Maximum = maxage
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        frm_perform_calculations.pbar.Value = 2
        Call delay(1#)
        .Refresh
    End With
End With

'Labelling the graph data window
pf_graph(doc).Label8.Caption = maxage
document(doc).Label8.Caption = maxage
'Label for t=pf(10%)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_pf(i, 2) <= 10# And doc_props(doc).graph_pf(i + 1, 2) >= 10#) Then
        mgraph = (doc_props(doc).graph_pf(i + 1, 2) - doc_props(doc).graph_pf(i, 2)) / (doc_props(doc).graph_pf(i + 1, 1) - doc_props(doc).graph_pf(i, 1))
        bgraph = doc_props(doc).graph_pf(i, 2) - mgraph * doc_props(doc).graph_pf(i, 1)
        'aux1 = 10 * mgraph + bgraph
        aux1 = (10 - bgraph) / mgraph
        pf_graph(doc).Label10.Caption = CStr(Round(aux1, 1))
        document(doc).Label10.Caption = CStr(Round(aux1, 1))
        aux2 = 123
        Exit For
    End If
Next i
If aux2 = 0 Then pf_graph(doc).Label10.Caption = "---"
If aux2 = 0 Then document(doc).Label10.Caption = "---"
'Label for t=pf(50%)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_pf(i, 2) <= 50# And doc_props(doc).graph_pf(i + 1, 2) >= 50#) Then
        mgraph = (doc_props(doc).graph_pf(i + 1, 2) - doc_props(doc).graph_pf(i, 2)) / (doc_props(doc).graph_pf(i + 1, 1) - doc_props(doc).graph_pf(i, 1))
        bgraph = doc_props(doc).graph_pf(i, 2) - mgraph * doc_props(doc).graph_pf(i, 1)
        'aux1 = 50 * mgraph + bgraph
        aux1 = (50 - bgraph) / mgraph
        pf_graph(doc).Label9.Caption = CStr(Round(aux1, 1))
        document(doc).Label9.Caption = CStr(Round(aux1, 1))
        aux2 = 123
        Exit For
    End If
Next i
If aux2 = 0 Then pf_graph(doc).Label9.Caption = "---"
If aux2 = 0 Then document(doc).Label9.Caption = "---"
'Label for t=pf(90%)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_pf(i, 2) < 90# And doc_props(doc).graph_pf(i + 1, 2) > 90#) Then
        mgraph = (doc_props(doc).graph_pf(i + 1, 2) - doc_props(doc).graph_pf(i, 2)) / (doc_props(doc).graph_pf(i + 1, 1) - doc_props(doc).graph_pf(i, 1))
        bgraph = doc_props(doc).graph_pf(i, 2) - mgraph * doc_props(doc).graph_pf(i, 1)
        'aux1 = 90 * mgraph + bgraph
        aux1 = (90 - bgraph) / mgraph
        pf_graph(doc).Label18.Caption = CStr(Round(aux1, 1))
        document(doc).Label18.Caption = CStr(Round(aux1, 1))
        aux2 = 123
        Exit For
    End If
Next i
If aux2 = 0 Then pf_graph(doc).Label18.Caption = "---"
If aux2 = 0 Then document(doc).Label18.Caption = "---"
'Label for t=pf(95%)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_pf(i, 2) < 95# And doc_props(doc).graph_pf(i + 1, 2) > 95#) Then
        mgraph = (doc_props(doc).graph_pf(i + 1, 2) - doc_props(doc).graph_pf(i, 2)) / (doc_props(doc).graph_pf(i + 1, 1) - doc_props(doc).graph_pf(i, 1))
        bgraph = doc_props(doc).graph_pf(i, 2) - mgraph * doc_props(doc).graph_pf(i, 1)
        'aux1 = 95 * mgraph + bgraph
        aux1 = (95 - bgraph) / mgraph
        pf_graph(doc).Label17.Caption = CStr(Round(aux1, 1))
        document(doc).Label17.Caption = CStr(Round(aux1, 1))
        aux2 = 123
        Exit For
    End If
Next i
If aux2 = 0 Then pf_graph(doc).Label17.Caption = "---"
If aux2 = 0 Then document(doc).Label17.Caption = "---"
'Label for t=pf(99%)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_pf(i, 2) < 99# And doc_props(doc).graph_pf(i + 1, 2) > 99#) Then
        mgraph = (doc_props(doc).graph_pf(i + 1, 2) - doc_props(doc).graph_pf(i, 2)) / (doc_props(doc).graph_pf(i + 1, 1) - doc_props(doc).graph_pf(i, 1))
        bgraph = doc_props(doc).graph_pf(i, 2) - mgraph * doc_props(doc).graph_pf(i, 1)
        aux1 = (99 - bgraph) / mgraph
        pf_graph(doc).Label16.Caption = CStr(Round(aux1, 1))
        document(doc).Label16.Caption = CStr(Round(aux1, 1))
        aux2 = 123
        Exit For
    End If
Next i
If aux2 = 0 Then pf_graph(doc).Label16.Caption = "---"
If aux2 = 0 Then document(doc).Label16.Caption = "---"

'************************************
'DRAWING GRAPH Reliability Index
'************************************

    'Calculating y values on graph
    betamax = Round(doc_props(doc).graph_beta(1, 2))
    If (betamax > 0 And betamax < doc_props(doc).graph_beta(1, 2)) Then betamax = betamax + 1
    If (betamax < 0 And betamax < doc_props(doc).graph_beta(1, 2)) Then betamax = betamax + 1
    betamin = Round(doc_props(doc).graph_beta(numdt, 2))
    If (betamin > 0 And betamin > doc_props(doc).graph_beta(numdt, 2)) Then betamin = betamin - 1
    If (betamin < 0 And betamin > doc_props(doc).graph_beta(numdt, 2)) Then betamin = betamin - 1

    With .reliability_chart
        .RowCount = numdt
        .ChartData = doc_props(doc).graph_beta()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = "Reliability Index"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.Maximum = betamax
            .ValueScale.Minimum = betamin
        End With

        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = "Time (years)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = False
            .AxisGrid.MajorPen.Style = VtPenStyleNull
            .AxisGrid.MinorPen.Style = VtPenStyleNull
            
            .ValueScale.MajorDivision = time_scale
            .ValueScale.MinorDivision = 0
            
            .ValueScale.Maximum = maxage
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        frm_perform_calculations.pbar.Value = 3
        Call delay(1#)
        .Refresh
    End With
With ri_graph(doc)
    With .reliability_chart
        .RowCount = numdt
        .ChartData = doc_props(doc).graph_beta()
        .chartType = VtChChartType2dXY
        ' Set Guide Lines for 2D XY chart Series 1.
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
        With .Plot.Axis(VtChAxisIdY)
            .AxisTitle = "Reliability Index"
            .AxisScale.Hide = False
            .CategoryScale.Auto = True
            .ValueScale.Maximum = betamax
            .ValueScale.Minimum = betamin
        End With

        With .Plot.Axis(VtChAxisIdX)
            .AxisTitle = "Time (years)"
            .AxisScale.Hide = False
            .CategoryScale.Auto = False
            .AxisGrid.MajorPen.Style = VtPenStyleNull
            .AxisGrid.MinorPen.Style = VtPenStyleNull
            
            .ValueScale.MajorDivision = time_scale
            .ValueScale.MinorDivision = 0
            
            .ValueScale.Maximum = maxage
            .ValueScale.Minimum = 0
        End With
        .Plot.UniformAxis = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = False
        .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdY) = False
        .Visible = True
        frm_perform_calculations.pbar.Value = 4
        Call delay(1#)
        .Refresh
    End With
End With

'Labelling the graph data window
ri_graph(doc).Label8.Caption = maxage
document(doc).Label4.Caption = maxage

'Label for t=beta(1.0)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_beta(i, 2) > 1# And doc_props(doc).graph_beta(i + 1, 2) < 1#) Then
        mgraph = (doc_props(doc).graph_beta(i + 1, 2) - doc_props(doc).graph_beta(i, 2)) / (doc_props(doc).graph_beta(i + 1, 1) - doc_props(doc).graph_beta(i, 1))
        bgraph = doc_props(doc).graph_beta(i, 2) - mgraph * doc_props(doc).graph_beta(i, 1)
        aux1 = (1 - bgraph) / mgraph
        ri_graph(doc).Label10.Caption = CStr(Round(aux1, 1))
        document(doc).Label12.Caption = CStr(Round(aux1, 1))
        aux2 = 123
    End If
Next i
If aux2 = 0 Then ri_graph(doc).Label10.Caption = "---"
If aux2 = 0 Then document(doc).Label12.Caption = "---"
'Label for t=beta(1.5)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_beta(i, 2) > 1.5 And doc_props(doc).graph_beta(i + 1, 2) < 1.5) Then
        mgraph = (doc_props(doc).graph_beta(i + 1, 2) - doc_props(doc).graph_beta(i, 2)) / (doc_props(doc).graph_beta(i + 1, 1) - doc_props(doc).graph_beta(i, 1))
        bgraph = doc_props(doc).graph_beta(i, 2) - mgraph * doc_props(doc).graph_beta(i, 1)
        aux1 = (1.5 - bgraph) / mgraph
        ri_graph(doc).Label9.Caption = CStr(Round(aux1, 1))
        document(doc).Label11.Caption = CStr(Round(aux1, 1))
        aux2 = 123
    End If
Next i
If aux2 = 0 Then ri_graph(doc).Label9.Caption = "---"
If aux2 = 0 Then document(doc).Label11.Caption = "---"
'Label for t=beta(2.0)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_beta(i, 2) > 2# And doc_props(doc).graph_beta(i + 1, 2) < 2#) Then
        mgraph = (doc_props(doc).graph_beta(i + 1, 2) - doc_props(doc).graph_beta(i, 2)) / (doc_props(doc).graph_beta(i + 1, 1) - doc_props(doc).graph_beta(i, 1))
        bgraph = doc_props(doc).graph_beta(i, 2) - mgraph * doc_props(doc).graph_beta(i, 1)
        aux1 = (2 - bgraph) / mgraph
        ri_graph(doc).Label18.Caption = CStr(Round(aux1, 1))
        document(doc).Label25.Caption = CStr(Round(aux1, 1))
        aux2 = 123
    End If
Next i
If aux2 = 0 Then ri_graph(doc).Label18.Caption = "---"
If aux2 = 0 Then document(doc).Label25.Caption = "---"
'Label for t=beta(3.0)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_beta(i, 2) > 3# And doc_props(doc).graph_beta(i + 1, 2) < 3#) Then
        mgraph = (doc_props(doc).graph_beta(i + 1, 2) - doc_props(doc).graph_beta(i, 2)) / (doc_props(doc).graph_beta(i + 1, 1) - doc_props(doc).graph_beta(i, 1))
        bgraph = doc_props(doc).graph_beta(i, 2) - mgraph * doc_props(doc).graph_beta(i, 1)
        aux1 = (3 - bgraph) / mgraph
        ri_graph(doc).Label17.Caption = CStr(Round(aux1, 1))
        document(doc).Label24.Caption = CStr(Round(aux1, 1))
        aux2 = 123
    End If
Next i
If aux2 = 0 Then ri_graph(doc).Label17.Caption = "---"
If aux2 = 0 Then document(doc).Label24.Caption = "---"
'Label for t=beta(4.0)
aux2 = 0
For i = 2 To numdt
    If (doc_props(doc).graph_beta(i, 2) > 4# And doc_props(doc).graph_beta(i + 1, 2) < 4#) Then
        mgraph = (doc_props(doc).graph_beta(i + 1, 2) - doc_props(doc).graph_beta(i, 2)) / (doc_props(doc).graph_beta(i + 1, 1) - doc_props(doc).graph_beta(i, 1))
        bgraph = doc_props(doc).graph_beta(i, 2) - mgraph * doc_props(doc).graph_beta(i, 1)
        aux1 = (4 - bgraph) / mgraph
        ri_graph(doc).Label16.Caption = CStr(Round(aux1, 1))
        document(doc).Label23.Caption = CStr(Round(aux1, 1))
        aux2 = 123
    End If
Next i
If aux2 = 0 Then ri_graph(doc).Label16.Caption = "---"
If aux2 = 0 Then document(doc).Label23.Caption = "---"


End With

frm_perform_calculations.Hide
Unload frm_perform_calculations

Me.mnu_results.Enabled = True
FState(doc).calculated = True
Screen.MousePointer = vbDefault

End Sub


Private Sub mnuFileClose_Click()
Call unload_document
Me.mnu_run.Enabled = False
Me.mnu_insert_data.Caption = "Insert Data"
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Bug.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub


Private Sub mnuFileSaveAs_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  
  ' get the current form index
  cur_doc = current_form()
If cur_doc = False Then
    Exit Sub
End If
    
    ' Set CancelError is True
    Dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    Dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialogs.Filter = "All Files (*.*)|*.*|Duracon Files" & "(*.drc)|*.drc"
    ' Specify default filter
    Dialogs.FilterIndex = 2
    ' set the working directory the application dir
    Dialogs.InitDir = App.path
    ' Display the save dialog box
    Dialogs.ShowSave
    If Err.Number <> 0 Then
      Exit Sub
    End If
    ' get the name file and the path
    name = GetFile(Dialogs.filename)
    path = GetPath(Dialogs.filename)
    Call save_file(name, path, cur_doc)
End Sub

Private Sub mnuFileSave_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  
  ' get the current form index
  cur_doc = current_form()
 If cur_doc = False Then
    Exit Sub
End If
  If FState(cur_doc).saved Then
    MsgBox "Document already Saved!", vbOK + vbCritical, "Duracon"
    Exit Sub
  End If
  If Not FState(cur_doc).newname Then
    ' Set CancelError is True
    Dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    Dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialogs.Filter = "All Files (*.*)|*.*|Duracon Files" & "(*.drc)|*.drc"
    ' Specify default filter
    Dialogs.FilterIndex = 2
    ' set the working directory the application dir
    Dialogs.InitDir = App.path
    ' Display the save dialog box
    Dialogs.ShowSave
    If Err.Number <> 0 Then
      Exit Sub
    End If
    ' get the name file and the path
    name = GetFile(Dialogs.filename)
    path = GetPath(Dialogs.filename)
    Call save_file(name, path, cur_doc)
    Exit Sub
  End If
  Call save_file(FState(cur_doc).name, FState(cur_doc).path, cur_doc)
End Sub

Private Sub mnuFileOpen_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  Dim arraycount As Integer
  Dim i As Integer
  Dim vali As Integer
  Dim vals As Single
  Dim old_doc As Integer
  Dim tmp As Integer
  Dim pos As Integer
  Dim r() As String
  
  
' Set CancelError is True
  Dialogs.CancelError = True
 On Error Resume Next
  ' Set flags
  Dialogs.Flags = cdlOFNHideReadOnly And cdlOFNAllowMultiselect
  ' Set filters
  Dialogs.Filter = "All Files (*.*)|*.*|Duracon Files" & _
  "(*.drc)|*.drc"
  ' Specify default filter
  Dialogs.FilterIndex = 2
  ' Display the open dialog box
  Dialogs.ShowOpen
  If Err.Number = 32755 Then ' cancel was selected
    Exit Sub
  End If
  ' get the name file and the path
  name = GetFile(Dialogs.filename)
  path = GetPath(Dialogs.filename)
   
  r() = Split(name, ".")
  If r(UBound(r)) <> "drc" Then
    MsgBox "Invalid File Type. You must select a Valid *.DRC file.", vbOK + vbCritical, "Duracon"
    Exit Sub
  End If
   
  ' creates a new Child Form  and get the current form index
  On Error Resume Next
  arraycount = UBound(document)
  If Err <> 0 Then
      cur_doc = FileNew()
  Else
    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    old_doc = current_form
    If old_doc <> -1 Then  'there are child forms open
      If FState(old_doc).values = False Then
        FState(old_doc).deleted = True
        FState(old_doc).Dirty = False
      End If
    End If
    cur_doc = FileNew()
  End If
  ' change to the selected directory
  ChDir path

  document(cur_doc).Caption = name
  FState(cur_doc).path = path
  FState(cur_doc).name = name
  FState(cur_doc).saved = True
  FState(cur_doc).newname = True
  FState(cur_doc).calculated = False
  FState(cur_doc).values = True
  
  
With doc_props(cur_doc)
    .tt = 99
  
    If name <> "" Then
        Open name For Input As #1
       
        Line Input #1, .nprojt
        Line Input #1, .descrip
        Input #1, .datepjt
        Input #1, .iseedv
        Input #1, .idifcoef
        Input #1, .tseriev
        Input #1, .prfsz
        Input #1, .prfpred
        
        For i = 0 To 7
            Input #1, vali
            .prmdistn(i) = vali
        Next i
    
        For i = 0 To 7
            Input #1, vals
            .prmvone(i) = vals
        Next i
    
        For i = 0 To 7
            Input #1, vals
            .prmvtwo(i) = vals
        Next i
    
        For i = 0 To 11
            Input #1, vals
            .prfvdp(i) = vals
        Next i
    
        For i = 0 To 11
            Input #1, vals
            .prfvcs(i) = vals
        Next i
        
        Input #1, .seed
        Input #1, .nsimul
        Input #1, .predage
    
        Close
    End If
    .frm_ca_board1_values.project_name = .nprojt
    .frm_ead_board_values.project_name = .nprojt
    .frm_ca_board1_values.Description = .descrip
    .frm_ead_board_values.Description = .descrip
    .frm_ca_board1_values.project_date = .datepjt
    .frm_ca_board1_values.project_date = .datepjt
    .frm_ca_board1_values.values = True
    .frm_ca_board1_values.ready = True
    If .idifcoef = 0 Then
        If .prmvone(6) = 28 Then
            .frm_ca_board1_values.cdc = -1 ' design value - EAD mode
            .frm_ca_board1_values.values = False
            .frm_ca_board1_values.ready = False
            .frm_ead_board_values.values = True
            .frm_ead_board_values.ready = True
        ElseIf .prmvone(6) = 63 Then
            .frm_ca_board1_values.cdc = 2
        End If
    ElseIf .idifcoef = 1 Then
            .frm_ca_board1_values.cdc = 1
    ElseIf .idifcoef = 2 Then
            .frm_ca_board1_values.cdc = 3
    End If
    If .prmvone(6) <> 28 And .prmvone(6) <> 63 Then
        .frm_ca_board1_values.testage_val = .prmvone(6)
    Else
        .frm_ca_board1_values.testage_val = "N/A"
    End If
    If .prmvone(7) = 21 Then
        .frm_ca_board1_values.testtemp_val = "N/A"
    Else
        .frm_ca_board1_values.testtemp_val = .prmvone(7)
    End If
    If .prmvtwo(5) = 50 Then
        .frm_ca_board1_values.Timeseries_1 = 0
    ElseIf .prmvtwo(5) = 75 Then
        .frm_ca_board1_values.Timeseries_1 = 1
    ElseIf .prmvtwo(5) = 100 Then
        .frm_ca_board1_values.Timeseries_1 = 2
    ElseIf .prmvtwo(5) = 125 Then
        .frm_ca_board1_values.Timeseries_1 = 3
    ElseIf .prmvtwo(5) = 150 Then
        .frm_ca_board1_values.Timeseries_1 = 4
    End If
        
    If .frm_ca_board1_values.cdc = 3 Then
        .frm_ca_board2_values.values = True
        .frm_ca_board2_values.ready = True
        If .predage = "" And .predage = 0 Then
            .frm_ca_board2_values.Age = "N/A"
        Else
            .frm_ca_board2_values.Age = .predage
        End If
        tmp = True
        For i = 11 To 0 Step -1
            If .prfvdp(i) <> 0 And tmp Then
                tmp = False
                pos = i
                Exit For
            End If
        Next i
        .frm_ca_board2_values.Prfsize = i - 5
    Else
        .frm_ca_board2_values.values = False
        .frm_ca_board2_values.ready = False
    End If
    
    For i = 0 To 4
     If .prmdistn(i) = 7 Then
        .frm_ca_board3_values.Distype(i) = 0
        .frm_ead_board_values.Distype(i) = 0
     ElseIf .prmdistn(i) = 0 Then
        .frm_ca_board3_values.Distype(i) = 1
        .frm_ead_board_values.Distype(i) = 1
     ElseIf .prmdistn(i) = 1 Then
        .frm_ca_board3_values.Distype(i) = 2
        .frm_ead_board_values.Distype(i) = 2
     ElseIf .prmdistn(i) = 2 Then
        .frm_ca_board3_values.Distype(i) = 3
        .frm_ead_board_values.Distype(i) = 3
     End If
    Next i
    .frm_ca_board3_values.ready = True
    .frm_ca_board3_values.values = True
End With

Call refresh_lista(cur_doc)

tmp = False
If Not (doc_props(cur_doc).frm_ca_board3_values.ready And doc_props(cur_doc).frm_ca_board1_values.ready) Then
    If doc_props(cur_doc).frm_ca_board1_values.cdc = 3 Then
        If doc_props(cur_doc).frm_ca_board2_values.ready Then
            tmp = False
        Else
            tmp = True
        End If
    Else
        tmp = True
    End If
End If

If tmp Then
    Exit Sub
End If
If FState(cur_doc).saved = False Then ' first you have to save the file
    Exit Sub
End If
Me.mnu_run.Enabled = True
Me.mnu_insert_data.Caption = "Edit Data"
End Sub

Private Sub mnuFileNew_Click()
    FileNew
    frm_choice.Show 1
End Sub

Private Sub Timer_Timer()
Dim doc As Integer
Dim tmp As Boolean
Dim arraycount As Integer
Dim i As Integer

Dim path As String
On Error Resume Next
arraycount = UBound(document)
If Err.Number <> 0 Then
    tmp = False
Else ' documents found - check if they really exist
    tmp = False
    For i = 1 To arraycount
        If Not FState(i).deleted Then
            tmp = True
            Exit For
        End If
    Next i
End If

If tmp Then
    Me.mnu_insert_data = True
    Me.mnu_clear_data = True
    Me.mnuFileClose.Enabled = True
Else
    Me.mnu_insert_data = False
    Me.mnu_clear_data = False
    Me.mnuFileClose.Enabled = False
End If

doc = current_form
If doc = -1 Then
    Exit Sub
End If
tmp = False
If doc_props(doc).frm_ead_board_values.ready Then
    tmp = False
ElseIf Not (doc_props(doc).frm_ca_board3_values.ready And doc_props(doc).frm_ca_board1_values.ready) Then
    If doc_props(doc).frm_ca_board1_values.cdc = 3 Then
        If doc_props(doc).frm_ca_board2_values.ready Then
            tmp = False
        Else
            tmp = True
        End If
    Else
        tmp = True
    End If
End If

If tmp Then
    Me.mnu_run.Enabled = False
    Me.mnu_insert_data.Caption = "Insert Data"
Else
    Me.mnu_run.Enabled = True
    Me.mnu_insert_data.Caption = "Edit Data"

End If
If FState(doc).calculated Then
    Me.mnu_results.Enabled = True
Else
    Me.mnu_results.Enabled = False
End If

End Sub

