VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "."
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
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
            TextSave        =   "16-06-2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:15"
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
         Caption         =   "&Insert Data"
      End
      Begin VB.Menu mnu_clear 
         Caption         =   "&Clear Data"
      End
   End
   Begin VB.Menu mnu_simulation 
      Caption         =   "&Simulation"
      Enabled         =   0   'False
      Begin VB.Menu mnu_run_analysis 
         Caption         =   "&Run Analysis"
      End
   End
   Begin VB.Menu mnu_results 
      Caption         =   "&Results"
      Enabled         =   0   'False
      Begin VB.Menu mnu_report 
         Caption         =   "&Analysis Report"
      End
      Begin VB.Menu mnu_graphics 
         Caption         =   "&Graphics"
         Begin VB.Menu mnu_global_analysis 
            Caption         =   "&Global analysis"
         End
         Begin VB.Menu mnu_structure_costs 
            Caption         =   "&Structure Costs"
         End
         Begin VB.Menu mnu_energy 
            Caption         =   "&Energy Consuption"
         End
         Begin VB.Menu mnu_water 
            Caption         =   "&Water Consuption"
         End
         Begin VB.Menu mnu_nox_emissions 
            Caption         =   "&NOx emissions"
         End
         Begin VB.Menu menu_co2_emissions 
            Caption         =   "&CO2 emissions"
         End
         Begin VB.Menu mnu_so2_emissions 
            Caption         =   "&SO2 emissions"
         End
      End
   End
   Begin VB.Menu mnu_database 
      Caption         =   "&Database"
      Begin VB.Menu mnu_database_maintenance 
         Caption         =   "&Open Database"
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
    Me.Caption = App.Title
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnu_clear_Click()
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
document(doc).lista.Clear
FState(doc).count = 0
    FState(doc).count = 0
    With document(doc)
        With .lista
            .Row = 0
            .Col = 0
            .ColWidth(0) = document(doc).TextWidth("#######") * 2
            .Refresh
            .Text = "Type"
            .ColWidth(1) = document(doc).TextWidth("####") * 2
            .Col = 1
            .CellAlignment = 4
            .Text = "Quantity"
            .ColWidth(2) = document(doc).TextWidth("####") * 2
            .Col = 2
            .CellAlignment = 4
            .Text = "Height"
            .ColWidth(3) = document(doc).TextWidth("####") * 2
            .Col = 3
            .CellAlignment = 4
            .Text = "Width"
            .ColWidth(4) = document(doc).TextWidth("####") * 2
            .Col = 4
            .CellAlignment = 4
            .Text = "Weight"
            .ColWidth(5) = document(doc).TextWidth("####") * 2
            .Col = 5
            .CellAlignment = 4
            .Text = "Lenght"
            .ColWidth(6) = document(doc).TextWidth("###") * 2
            .Col = 6
            .CellAlignment = 4
            .Text = "f5"
            .ColWidth(7) = document(doc).TextWidth("###") * 2
            .Col = 7
            .CellAlignment = 4
            .Text = "f6"
            .ColWidth(8) = document(doc).TextWidth("###") * 2
            .Col = 8
            .CellAlignment = 4
            .Text = "f8"
            .ColWidth(9) = document(doc).TextWidth("###") * 2
            .Col = 9
            .CellAlignment = 4
            .Text = "f10"
            .ColWidth(10) = document(doc).TextWidth("###") * 2
            .Col = 10
            .CellAlignment = 4
            .Text = "f12"
            .ColWidth(11) = document(doc).TextWidth("###") * 2
            .Col = 11
            .CellAlignment = 4
            .Text = "f16"
            .ColWidth(12) = document(doc).TextWidth("###") * 2
            .Col = 12
            .CellAlignment = 4
            .Text = "f20"
            .ColWidth(13) = document(doc).TextWidth("###") * 2
            .Col = 13
            .CellAlignment = 4
            .Text = "f25"
            .ColWidth(14) = document(doc).TextWidth("###") * 2
            .Col = 14
            .CellAlignment = 4
            .Text = "f32"
            .ColWidth(15) = document(doc).TextWidth("####") * 2
            .Col = 15
            .CellAlignment = 4
            .Text = "Cement"
            .ColWidth(16) = document(doc).TextWidth("#####") * 2
            .Col = 16
            .CellAlignment = 4
            .Text = "Aggregates"
            .ColWidth(17) = document(doc).TextWidth("####") * 2
            .Col = 17
            .CellAlignment = 4
            .Text = "Costs"
            .ColWidth(18) = document(doc).TextWidth("######") * 2
            .Col = 18
            .CellAlignment = 4
            .Text = "Database"
        End With
    End With

End Sub

Private Sub mnu_database_maintenance_Click()
Me.mnuFile.Enabled = False
    Me.mnu_simulation.Enabled = False
    Me.mnuWindow.Enabled = False
    Me.mnu_graphics.Enabled = False
    Me.mnuEdit.Enabled = False
If Me.mnu_database_maintenance.Caption = "&Open Database" Then
    Me.mnu_database_maintenance.Caption = "&Close Database"
    Load Frm_database
    Frm_database.Show
Else
    Me.mnu_database_maintenance.Caption = "&Open Database"
    Unload Frm_database
End If
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
frm_choice.Show 1
End Sub

Private Sub mnu_energy_Click()
Dim doc As Integer

doc = current_form
frm_energy(doc).Show
End Sub

Private Sub mnu_global_analysis_Click()
Dim doc As Integer

doc = current_form
frm_global(doc).Show 1

End Sub

Private Sub mnu_nox_emissions_Click()
Dim doc As Integer

doc = current_form
frm_nox(doc).Show 1

End Sub

Private Sub mnu_report_Click()
Dim doc As Integer

doc = current_form
frm_report(doc).Show 1


End Sub



Private Sub mnu_run_analysis_Click()

Dim doc As Integer
Dim i As Integer
Dim peso, peso_total As Double
Dim impact_metal As impact
Dim impact_concrete As impact
Dim db_pos As Integer
Dim filename As String
Dim num As Integer
Dim chain As String
Dim r() As String
Dim s() As String
Dim concrete() As concrete_type
Dim metalic() As steel_type
Dim qtd_cimento, qtd_agregados, qtd_armadura As Double
Dim vol_betao, vol_total_betao, vol_madeira, vol_total_madeira As Double
Dim tmp, lenght, qtd As Double
Dim steel_bars(9) As Double

Screen.MousePointer = vbHourglass
Load frm_perform_calculations
frm_perform_calculations.Show
frm_perform_calculations.SetFocus
Call DisableX(frm_perform_calculations)
Call delay(0.02)
frm_perform_calculations.txt.Caption = "Running analysis..."

doc = current_form
With document(doc)
    ' loading metallic struct database
    ReDim metalic(1)
    Err.Clear
    On Error Resume Next
    filename = App.path & "\database\steel.dbs"
    Open filename For Input As #1
    Input #1, num
    ReDim metalic(num + 1)
    i = 0
    While Not EOF(1)
        Input #1, chain
        i = i + 1
        r() = Split(chain, "@")
        s() = Split(r(0), "#")
        With metalic(i)
            .name = s(0)
            .date = s(1)
            .description = s(2)
            s() = Split(r(1), "#")
            With .steel
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(2), "#")
            With .transport
                .co2 = str2str(s(0))
                .distance = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
            End With
        End With
    Wend
    Close #1
        
    peso_total = 0
    db_pos = -1
    For i = 1 To FState(doc).count
        .lista.Row = i
        .lista.Col = 0
        If db_pos = -1 Then
            db_pos = i
        End If
        If .lista.Text = "Met.Beam" Or .lista.Text = "Met.Pillar" Then
            .lista.Col = 1
            peso = convert_type(.lista.Text) ' qtd
            .lista.Col = 4
            peso = peso * convert_type(.lista.Text) ' weight kg/m3
            .lista.Col = 5
            peso = peso * convert_type(.lista.Text) ' lenght
            peso_total = peso_total + peso
        End If
    Next i
    If peso_total <> 0 Then
        peso_total = peso_total / 1000 ' conversion kg to ton
        With impact_metal
            document(doc).lista.Row = db_pos
            document(doc).lista.Col = 17
            doc_props(doc).metal_cost = document(doc).lista.Text
            .costs = peso_total * 1000 * convert_type(document(doc).lista.Text)
            document(doc).lista.Row = db_pos
            document(doc).lista.Col = 19
            db_pos = convert_type(document(doc).lista.Text)
            doc_props(doc).impact_transport.co2 = Round((peso_total * metalic(db_pos).transport.co2 * metalic(db_pos).transport.distance / 1000) / 100, 2)
            doc_props(doc).impact_transport.so2 = Round(peso_total * metalic(db_pos).transport.so2 * metalic(db_pos).transport.distance / 1000, 2)
            doc_props(doc).impact_transport.nox = Round(peso_total * metalic(db_pos).transport.nox * metalic(db_pos).transport.distance / 1000, 2)
            .co2 = (peso_total * metalic(db_pos).steel.co2 * 1000 + peso_total * metalic(db_pos).transport.co2 * metalic(db_pos).transport.distance / 1000) / 100
            .energy = peso_total * metalic(db_pos).steel.energy
            .nox = peso_total * metalic(db_pos).steel.nox * 1000 + peso_total * metalic(db_pos).transport.nox * metalic(db_pos).transport.distance / 1000
            .so2 = peso_total * metalic(db_pos).steel.so2 * 1000 + peso_total * metalic(db_pos).transport.so2 * metalic(db_pos).transport.distance / 1000
            .water = peso_total * metalic(db_pos).steel.water
        End With
    End If
    With doc_props(doc)
        .total_weight = Round(peso_total, 2)
        With .impact_metal
            .co2 = Round(impact_metal.co2, 2)
            .costs = Round(impact_metal.costs, 2)
            .energy = Round(impact_metal.energy, 2)
            .nox = Round(impact_metal.nox, 2)
            .so2 = Round(impact_metal.so2, 2)
            .water = Round(impact_metal.water, 2)
        End With
    End With
    ' Calculating Concrete strutures
    ReDim concrete(1)
    Err.Clear
    On Error Resume Next
    filename = App.path & "\database\concrete.dbs"
    Open filename For Input As #1
    Input #1, num
    ReDim concrete(num + 1)
    i = 0
    While Not EOF(1)
        Input #1, chain
        i = i + 1
        r() = Split(chain, "@")
        s() = Split(r(0), "#")
        With concrete(i)
            .name = s(0)
            .date = s(1)
            .description = s(2)
            s() = Split(r(1), "#")
            With .wood
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(2), "#")
            With .cement
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(3), "#")
            With .steel
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(4), "#")
            With .agregates
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(5), "#")
            With .water
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
        End With
    Wend
    Close #1
    vol_betao = 0
    vol_madeira = 0
    vol_total_betao = 0
    vol_total_madeira = 0
    For i = 1 To 9
        steel_bars(i) = 0
    Next i
    db_pos = -1
    For i = 1 To FState(doc).count
        .lista.Row = i
        .lista.Col = 0
        If .lista.Text = "Conc.Beam" Or .lista.Text = "Conc.Pillar" Then
            If db_pos = -1 Then
                db_pos = i
            End If
            .lista.Col = 1
            vol_betao = convert_type(.lista.Text) ' qtd
            vol_madeira = 0.03 * 2 * convert_type(.lista.Text) ' qtd
            .lista.Col = 5
            vol_betao = vol_betao * convert_type(.lista.Text) ' lenght m
            vol_madeira = vol_madeira * convert_type(.lista.Text) ' lenght m
            .lista.Col = 3
            vol_betao = vol_betao * convert_type(.lista.Text) ' width m
            tmp = convert_type(.lista.Text) ' width m
            .lista.Col = 2
            vol_betao = vol_betao * convert_type(.lista.Text) ' height m
            tmp = tmp + convert_type(.lista.Text) ' height m
            vol_madeira = vol_madeira * tmp
            
            vol_total_madeira = vol_total_madeira + vol_madeira
            vol_total_betao = vol_total_betao + vol_betao
            .lista.Col = 5
            lenght = convert_type(.lista.Text)
            .lista.Col = 1
            qtd = convert_type(.lista.Text)
            .lista.Col = 6
            steel_bars(1) = steel_bars(1) + lenght * qtd * convert_type(.lista.Text) ' varao de 5
            .lista.Col = 7
            steel_bars(2) = steel_bars(2) + lenght * qtd * convert_type(.lista.Text)  ' varao de 6
            .lista.Col = 8
            steel_bars(3) = steel_bars(3) + lenght * qtd * convert_type(.lista.Text)  ' varao de 8
            .lista.Col = 9
            steel_bars(4) = steel_bars(4) + lenght * qtd * convert_type(.lista.Text)  ' varao de 10
            .lista.Col = 10
            steel_bars(5) = steel_bars(5) + lenght * qtd * convert_type(.lista.Text)  ' varao de 12
            .lista.Col = 11
            steel_bars(6) = steel_bars(6) + lenght * qtd * convert_type(.lista.Text)  ' varao de 16
            .lista.Col = 12
            steel_bars(7) = steel_bars(7) + lenght * qtd * convert_type(.lista.Text)  ' varao de 20
            .lista.Col = 13
            steel_bars(8) = steel_bars(8) + lenght * qtd * convert_type(.lista.Text)  ' varao de 25
            .lista.Col = 14
            steel_bars(9) = steel_bars(9) + lenght * qtd * convert_type(.lista.Text)  ' varao de 32
       End If
    Next i
    .lista.Row = db_pos
    .lista.Col = 17
    doc_props(doc).concrete_cost = document(doc).lista.Text
    .lista.Col = 16
    doc_props(doc).aggregates = document(doc).lista.Text
    .lista.Col = 15
    doc_props(doc).cement = document(doc).lista.Text
    .lista.Col = 19
    db_pos = convert_type(.lista.Text)
    .lista.Col = 15
    qtd_cimento = vol_total_betao * convert_type(.lista.Text)
    .lista.Col = 16
    qtd_agregados = vol_total_betao * convert_type(.lista.Text)
    qtd_armadura = 0.151 * 1.05 * steel_bars(1) + 0.22 + 1.05 * steel_bars(2) + 0.395 * 1.05 * steel_bars(3) + 0.617 * 1.07 * steel_bars(4) + 0.888 * 1.07 * steel_bars(5) + 1.58 * 1.1 * steel_bars(6) + 2.47 * 1.1 * steel_bars(7) + 3.86 * 1.12 * steel_bars(8) + 6.31 * 1.12 * steel_bars(9)
    With impact_concrete
        .energy = concrete(db_pos).cement.energy * qtd_cimento / 1000 + concrete(db_pos).agregates.energy * qtd_agregados / 1000 + concrete(db_pos).water.energy * vol_total_betao + concrete(db_pos).steel.energy * qtd_armadura / 1000 + concrete(db_pos).wood.energy * vol_total_madeira
        .co2 = (concrete(db_pos).cement.co2 * qtd_cimento / 1000 + concrete(db_pos).agregates.co2 * qtd_agregados / 1000 + concrete(db_pos).water.co2 * vol_total_betao + concrete(db_pos).steel.co2 * qtd_armadura + concrete(db_pos).wood.co2 * vol_total_madeira * 1000) / 100
        .so2 = concrete(db_pos).cement.so2 * qtd_cimento / 1000 + concrete(db_pos).agregates.so2 * qtd_agregados / 1000 + concrete(db_pos).water.so2 * vol_total_betao + concrete(db_pos).steel.so2 * qtd_armadura + concrete(db_pos).wood.so2 * vol_total_madeira * 1000
        .nox = concrete(db_pos).cement.nox * qtd_cimento / 1000 + concrete(db_pos).agregates.nox * qtd_agregados / 1000 + concrete(db_pos).water.nox * vol_total_betao + concrete(db_pos).steel.nox * qtd_armadura + concrete(db_pos).wood.nox * vol_total_madeira * 1000
        .water = concrete(db_pos).cement.water * qtd_cimento / 1000 + concrete(db_pos).agregates.water * qtd_agregados / 1000 + concrete(db_pos).water.water * vol_total_betao + concrete(db_pos).steel.water * qtd_armadura / 1000 + concrete(db_pos).wood.water * vol_total_madeira
        document(doc).lista.Col = 17
        .costs = vol_total_betao * convert_type(document(doc).lista.Text)
    End With
    With doc_props(doc)
        .armour_qty = Round(qtd_armadura, 2)
        .aggregates_qty = Round(qtd_agregados, 2)
        .cement_qty = Round(qtd_cimento, 2)
        .volume_concrete = Round(vol_total_betao, 2)
        .volume_wood = Round(vol_total_madeira, 2)
        With .impact_concrete
            .co2 = Round(impact_concrete.co2, 2)
            .costs = Round(impact_concrete.costs, 2)
            .energy = Round(impact_concrete.energy, 2)
            .nox = Round(impact_concrete.nox, 2)
            .so2 = Round(impact_concrete.so2, 2)
            .water = Round(impact_concrete.water, 2)
        End With
        .impact_total.co2 = .impact_metal.co2 + .impact_concrete.co2
        .impact_total.costs = .impact_metal.costs + .impact_concrete.costs
        .impact_total.energy = .impact_metal.energy + .impact_concrete.energy
        .impact_total.nox = .impact_metal.nox + .impact_concrete.nox
        .impact_total.so2 = .impact_metal.so2 + .impact_concrete.so2
        .impact_total.water = .impact_metal.water + .impact_concrete.water
    End With
End With
frm_perform_calculations.pbar.Max = 6

With doc_props(doc).dados
    .costs(1, 1) = Round(impact_concrete.costs, 2)
    .costs(1, 2) = Round(impact_metal.costs, 2)
End With
frm_perform_calculations.txt.Caption = "Plotting Structure charts..."
Call draw_graph(document(doc).structure_chart, doc_props(doc).dados.costs, "", "€")
frm_perform_calculations.pbar.Value = 1

With doc_props(doc).dados
    .energy(1, 1) = Round(impact_concrete.energy, 2)
    .energy(1, 2) = Round(impact_metal.energy, 2)
End With
frm_perform_calculations.txt.Caption = "Plotting energy charts..."
Call draw_graph(document(doc).energy_chart, doc_props(doc).dados.energy, "", "GJ")
frm_perform_calculations.pbar.Value = 2

With doc_props(doc).dados
    .water(1, 1) = Round(impact_concrete.water, 2)
    .water(1, 2) = Round(impact_metal.water, 2)
End With
frm_perform_calculations.txt.Caption = "Plotting water charts..."
Call draw_graph(document(doc).water_chart, doc_props(doc).dados.water, "", "m3")
frm_perform_calculations.pbar.Value = 3

With doc_props(doc).dados
    .nox(1, 1) = Round(impact_concrete.nox, 2)
    .nox(1, 2) = Round(impact_metal.nox, 2)
End With
frm_perform_calculations.txt.Caption = "Plotting NOx charts..."
Call draw_graph(document(doc).nox_chart, doc_props(doc).dados.nox, "", "Kg")
frm_perform_calculations.pbar.Value = 4

With doc_props(doc).dados
    .so2(1, 1) = Round(impact_concrete.so2, 2)
    .so2(1, 2) = Round(impact_metal.so2, 2)
End With
frm_perform_calculations.txt.Caption = "Plotting SO2 charts..."
Call draw_graph(document(doc).so2_chart, doc_props(doc).dados.so2, "", "Kg")
frm_perform_calculations.pbar.Value = 5

With doc_props(doc).dados
    .co2(1, 1) = Round(impact_concrete.co2, 2)
    .co2(1, 2) = Round(impact_metal.co2, 2)
End With
frm_perform_calculations.txt.Caption = "Plotting CO2 charts..."
Call draw_graph(document(doc).co2_chart, doc_props(doc).dados.co2, "", "x100 Kg")
frm_perform_calculations.pbar.Value = 6

With doc_props(doc).dados
    .total(1, 1) = Round(impact_concrete.co2 + impact_concrete.so2 + impact_concrete.nox + impact_concrete.energy + impact_concrete.water, 2)
    .total(1, 2) = Round(impact_metal.co2 + impact_metal.so2 + impact_metal.nox + impact_metal.energy + impact_metal.water, 2)
End With

frm_perform_calculations.Hide
Unload frm_perform_calculations

FState(doc).calculated = True
Screen.MousePointer = vbDefault
End Sub
Private Sub draw_graph(grafico As MSChart, dado() As Double, x_title As String, y_title As String)
With grafico
    .RowCount = 1
    .ColumnCount = 2
    .ChartData = dado
    .chartType = VtChChartType3dBar
    .Row = 1
    .RowLabelIndex = 1
    .RowLabel = ""
    
    .Column = 1
    .ColumnLabelIndex = 1
    .ColumnLabel = "Concrete"
    .Column = 2
    .ColumnLabelIndex = 1
    .ColumnLabel = "Metallic"
    
    .Plot.View3d.Rotation = 0
    .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
    With .Plot.SeriesCollection(1)
        .DataPoints(-1).DataPointLabel.VtFont.name = "Verdana"
        .DataPoints(-1).DataPointLabel.VtFont.Size = 14
        .DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeRight
    End With
    With .Plot.SeriesCollection(2)
        .DataPoints(-1).DataPointLabel.VtFont.name = "Verdana"
        .DataPoints(-1).DataPointLabel.VtFont.Size = 14
        .DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeRight
    End With
    
    With .Plot.Axis(VtChAxisIdY)
        .AxisTitle = ""
        .AxisScale.Hide = False
        .CategoryScale.Auto = True
    End With

    With .Plot.Axis(VtChAxisIdX)
        .AxisTitle = ""
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



Private Sub mnu_so2_emissions_Click()
Dim doc As Integer

doc = current_form
frm_so2(doc).Show 1

End Sub

Private Sub mnu_structure_costs_Click()
Dim doc As Integer

doc = current_form
frm_structure(doc).Show 1

End Sub

Private Sub mnu_water_Click()
Dim doc As Integer

doc = current_form
frm_water(doc).Show 1
End Sub
Private Sub menu_co2_emissions_Click()
Dim doc As Integer

doc = current_form
frm_co2(doc).Show 1
End Sub

Private Sub mnuFileClose_Click()
Call unload_document
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
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
            MsgBox Err.description
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
 If FState(cur_doc).deleted Then
       MsgBox "This message Should not appear! Look into IT PLEASE!", vbOKCancel, "Info"
     Exit Sub
  End If
    
    ' Set CancelError is True
    Dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    Dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialogs.Filter = dialogs_filter
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
If FState(cur_doc).deleted Then
       MsgBox "ERROR -This message Should not appear!", vbOKCancel, "Info"
     Exit Sub
  End If
  If FState(cur_doc).saved Then
    MsgBox "Document already Saved!", vbOKCancel, "Info"
    Exit Sub
  End If
  If Not FState(cur_doc).newname Then
    ' Set CancelError is True
    Dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    Dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialogs.Filter = dialogs_filter
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
  Dim old_doc As Integer
  Dim tmp As Integer
  Dim r() As String
  Dim chain As String
  Dim filename As String
  
' Set CancelError is True
  Dialogs.CancelError = True
 On Error Resume Next
  ' Set flags
  Dialogs.Flags = cdlOFNHideReadOnly And cdlOFNAllowMultiselect
  ' Set filters
  Dialogs.Filter = dialogs_filter
  ' Specify default filter
  Dialogs.FilterIndex = 2
  ' Display the open dialog box
  Dialogs.ShowOpen
  If Err.Number = 32755 Then ' cancel was selected
    'MsgBox "error num:" & Str(Err.Number) & " Desc:" & Err.Description, vbOKCancel, "Info"
    Exit Sub
  End If
  ' get the name file and the path
  name = GetFile(Dialogs.filename)
  path = GetPath(Dialogs.filename)
   
  r() = Split(name, ".")
  If r(UBound(r)) <> filename_extension Then
    MsgBox "Invalid File Type. You must select a Valid *." & filename_extension & " file.", vbOK, "Info"
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
filename = name
Open filename For Input As #1
Line Input #1, chain
FState(cur_doc).count = 0
document(cur_doc).lista.Visible = False
While Not EOF(1)
    Line Input #1, chain
    FState(cur_doc).count = FState(cur_doc).count + 1
    With document(cur_doc).lista
        If .Rows <= FState(cur_doc).count Then
            .Rows = .Rows + 10
        End If
        r() = Split(chain, ";")
        .Row = FState(cur_doc).count
        .Col = 0
        .Text = r(0)
        For i = 1 To 19
            .Col = i
            .Text = str2str(r(i))
        Next i
    End With
    
Wend
document(cur_doc).lista.Visible = True
Close #1


  document(cur_doc).Caption = name
  FState(cur_doc).path = path
  FState(cur_doc).name = name
  FState(cur_doc).saved = True
  FState(cur_doc).newname = True
  FState(cur_doc).calculated = False
  FState(cur_doc).values = True
  
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

If tmp Then ' documents found
    If Me.mnu_simulation.Enabled = False Then
        Me.mnu_simulation.Enabled = True
    End If
    doc = current_form
    If FState(doc).calculated = True Then
        If Me.mnu_results.Enabled = False Then
            Me.mnu_results.Enabled = True
        End If
    End If
Else
    If Me.mnu_simulation.Enabled = True Then
        Me.mnu_simulation.Enabled = False
    End If
End If

If check_db = True Then
    Me.mnu_database_maintenance.Caption = "&Open Database"
    Me.mnuFile.Enabled = True
    Me.mnu_simulation.Enabled = True
    Me.mnuWindow.Enabled = True
    Me.mnu_graphics.Enabled = True
    Me.mnuEdit.Enabled = True
    check_db = False
End If
End Sub

