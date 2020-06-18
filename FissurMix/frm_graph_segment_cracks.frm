VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_graph_segment_cracks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[CODED]"
   ClientHeight    =   8130
   ClientLeft      =   1995
   ClientTop       =   2610
   ClientWidth     =   11370
   Icon            =   "frm_graph_segment_cracks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleMode       =   0  'User
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
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
      Begin MSChart20Lib.MSChart chart_exp_data 
         Height          =   7605
         Left            =   120
         OleObjectBlob   =   "frm_graph_segment_cracks.frx":08CA
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   11115
      End
      Begin VB.Label segments_txt 
         Caption         =   "[CODED]"
         Height          =   525
         Left            =   210
         TabIndex        =   1
         Top             =   30
         Width           =   2805
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
Attribute VB_Name = "frm_graph_segment_cracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim doc As Integer
Dim vector() As Double
Dim i, contador As Integer
Dim label_x As String
Dim label_y As String
Dim tmp, tmp2 As Double
Dim num_segments As Integer

doc = current_form
Call DisableX(frm_exp_data(doc))

Me.Caption = "Cracks per segment Graph"
num_segments = 30

segments_txt.Caption = "Number of segments: " & CStr(num_segments)
With doc_props(doc)
      tmp = .statistic.elements / num_segments ' element per segment
      ReDim vector(1 To num_segments, 1)
      tmp2 = tmp
      contador = 1
      For i = 1 To .statistic.elements
        If i = tmp2 + 1 Then
            contador = contador + 1
            tmp2 = tmp2 + tmp
        End If
        If .elements(i).cracked Then
            vector(contador, 1) = vector(contador, 1) + 1
        End If
      Next i
      
End With

label_x = "Segments"
label_y = "Cracks"
With frm_segment_cracks(doc).chart_exp_data
    .ChartData = vector()
    .chartType = VtChChartType2dBar
        
    .Plot.SeriesCollection.Item(1).ShowGuideLine(VtChAxisIdX) = True
    With .Plot.SeriesCollection(1)
        .DataPoints(-1).DataPointLabel.VtFont.name = "Verdana"
        .DataPoints(-1).DataPointLabel.VtFont.Size = 10
        '.DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeRight
    End With
    With .Plot.SeriesCollection(2)
        .DataPoints(-1).DataPointLabel.VtFont.name = "Verdana"
        .DataPoints(-1).DataPointLabel.VtFont.Size = 10
        '.DataPoints.Item(-1).DataPointLabel.LocationType = VtChLabelLocationTypeRight
    End With
    
    With .Plot.Axis(VtChAxisIdY)
        .AxisTitle = label_y
        .AxisScale.Hide = False
        .CategoryScale.Auto = True
    End With

    With .Plot.Axis(VtChAxisIdX)
        .AxisTitle = label_x
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

frm_segment_cracks(doc).chart_exp_data.EditCopy
End Sub

Private Sub exit_Click()
Dim doc As Integer
doc = current_form

frm_segment_cracks(doc).Hide
Unload Me
End Sub


Private Sub printgraph_Click()
Dim doc As Integer
doc = current_form

frm_segment_cracks(doc).PrintForm
End Sub

