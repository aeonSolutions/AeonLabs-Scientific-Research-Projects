VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_graph_energy 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Energy Consuption Graph"
   ClientHeight    =   8130
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   11370
   Icon            =   "frm_graph_energy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleMode       =   0  'User
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   8055
      Index           =   0
      Left            =   0
      ScaleHeight     =   8025
      ScaleWidth      =   11265
      TabIndex        =   0
      Top             =   0
      Width           =   11295
      Begin MSChart20Lib.MSChart chart_energy 
         Height          =   7905
         Left            =   0
         OleObjectBlob   =   "frm_graph_energy.frx":2052
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   11115
      End
   End
   Begin VB.Menu viewother 
      Caption         =   "&View other graphs"
      Begin VB.Menu mnu_global 
         Caption         =   "Global Analysis"
      End
      Begin VB.Menu mnu_energy 
         Caption         =   "Structure Costs"
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
         Caption         =   "SO2 emissions"
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
Attribute VB_Name = "frm_graph_energy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim doc As Integer

doc = current_form
Call DisableX(frm_energy(doc))

With frm_energy(doc).chart_energy
    .RowCount = 1
    .ColumnCount = 2
    .ChartData = doc_props(doc).dados.energy
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
        .AxisTitle = "GJ"
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



Private Sub CopyGraph_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).chart_energy.EditCopy
End Sub

Private Sub exit_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).Hide
Unload Me
End Sub


Private Sub mnu_co2_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).Hide
Unload Me
frm_co2(doc).Show 1


End Sub

Private Sub mnu_energy_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).Hide
Unload Me
frm_structure(doc).Show 1

End Sub

Private Sub mnu_global_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).Hide
Unload Me
frm_global(doc).Show 1

End Sub

Private Sub mnu_nox_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).Hide
Unload Me
frm_nox(doc).Show 1

End Sub

Private Sub mnu_so2_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).Hide
Unload Me
frm_so2(doc).Show 1

End Sub

Private Sub mnu_water_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).Hide
Unload Me
frm_water(doc).Show 1

End Sub

Private Sub printgraph_Click()
Dim doc As Integer
doc = current_form

frm_energy(doc).PrintForm
End Sub

