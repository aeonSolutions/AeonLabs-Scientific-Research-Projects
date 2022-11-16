VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDocument 
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab 
      Height          =   10455
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   18441
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Input / Output Data"
      TabPicture(0)   =   "frmDocument.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "results"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lista"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Temperature Graph"
      TabPicture(1)   =   "frmDocument.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "temp_rotation_Scroll"
      Tab(1).Control(1)=   "chart_type"
      Tab(1).Control(2)=   "temperature_big_chart"
      Tab(1).Control(3)=   "temp_angle"
      Tab(1).Control(4)=   "temp_rotation_txt_angle"
      Tab(1).Control(5)=   "temp_rotation_txt_info"
      Tab(1).Control(6)=   "Label1"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Displacement Graph"
      TabPicture(2)   =   "frmDocument.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "displacement_big_chart"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Stress Graph"
      TabPicture(3)   =   "frmDocument.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "tension_big_chart"
      Tab(3).ControlCount=   1
      Begin VB.HScrollBar temp_rotation_Scroll 
         Height          =   285
         Left            =   -72810
         Max             =   360
         TabIndex        =   15
         Top             =   9390
         Width           =   11115
      End
      Begin VB.ComboBox chart_type 
         Height          =   315
         Left            =   -73710
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   600
         Width           =   1905
      End
      Begin MSChart20Lib.MSChart temperature_big_chart 
         Height          =   8385
         Left            =   -74250
         OleObjectBlob   =   "frmDocument.frx":093A
         TabIndex        =   9
         Top             =   960
         Width           =   12975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Temperature Gradient"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3645
         Left            =   150
         TabIndex        =   7
         Top             =   6480
         Width           =   4575
         Begin MSChart20Lib.MSChart temperature_chart 
            Height          =   3405
            Left            =   30
            OleObjectBlob   =   "frmDocument.frx":2C8A
            TabIndex        =   8
            Top             =   210
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Displacement Gradient"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3645
         Left            =   5280
         TabIndex        =   6
         Top             =   6480
         Width           =   4575
         Begin MSChart20Lib.MSChart displacement_chart 
            Height          =   3345
            Left            =   60
            OleObjectBlob   =   "frmDocument.frx":5193
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stress Gradient"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3645
         Left            =   10380
         TabIndex        =   5
         Top             =   6480
         Width           =   4575
         Begin MSChart20Lib.MSChart tension_chart 
            Height          =   3375
            Left            =   30
            OleObjectBlob   =   "frmDocument.frx":769C
            TabIndex        =   19
            Top             =   210
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   1035
         Left            =   150
         TabIndex        =   1
         Top             =   600
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   1826
         _Version        =   393216
         Rows            =   6
         Cols            =   10
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid results 
         Height          =   4245
         Left            =   150
         TabIndex        =   3
         Top             =   2040
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   7488
         _Version        =   393216
         Rows            =   100
         Cols            =   20
         FixedCols       =   0
         ScrollTrack     =   -1  'True
      End
      Begin MSChart20Lib.MSChart displacement_big_chart 
         Height          =   8385
         Left            =   -74640
         OleObjectBlob   =   "frmDocument.frx":9BA5
         TabIndex        =   10
         Top             =   600
         Width           =   12975
      End
      Begin MSChart20Lib.MSChart tension_big_chart 
         Height          =   8385
         Left            =   -73290
         OleObjectBlob   =   "frmDocument.frx":BEF5
         TabIndex        =   11
         Top             =   1020
         Width           =   12975
      End
      Begin VB.Label temp_angle 
         Caption         =   "rotation angle:"
         Height          =   315
         Left            =   -68760
         TabIndex        =   17
         Top             =   9780
         Width           =   1365
      End
      Begin VB.Label temp_rotation_txt_angle 
         Caption         =   "0º"
         Height          =   255
         Left            =   -67410
         TabIndex        =   16
         Top             =   9780
         Width           =   795
      End
      Begin VB.Label temp_rotation_txt_info 
         Caption         =   "Rotation 3D Graph :"
         Height          =   285
         Left            =   -74730
         TabIndex        =   14
         Top             =   9420
         Width           =   1815
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Left            =   -74850
         TabIndex        =   12
         Top             =   630
         Width           =   1125
         Caption         =   "Chart Type:"
         Size            =   "1984;450"
         FontName        =   "Verdana"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label12 
         Caption         =   "Results"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   4
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label11 
         Caption         =   "Material propreties"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog dialogs 
      Left            =   5280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4560
      Top             =   0
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chart_type_click()
    Select Case chart_type.ListIndex
    Case 0 To 9
        temperature_big_chart.chartType = chart_type.ListIndex
    Case 10
        temperature_big_chart.chartType = VtChChartType2dPie
    Case 11
        temperature_big_chart.chartType = VtChChartType2dXY
    End Select
    If temperature_big_chart.Chart3d = True Then
        temp_rotation_txt_info.Enabled = True
        temp_rotation_Scroll.Enabled = True
        temp_rotation_txt_angle.Enabled = True
        temp_angle.Enabled = True
    Else
        temp_rotation_txt_info.Enabled = False
        temp_rotation_Scroll.Enabled = False
        temp_rotation_txt_angle.Enabled = False
        temp_angle.Enabled = False
    End If

End Sub

Sub form_activate()
    Dim arraycount As Integer
    Dim i As Integer
    
    arraycount = UBound(document)

    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    FState(Me.Tag).Dirty = True
End Sub
Sub form_deactivate()
    Dim arraycount As Integer
    Dim i As Integer
    arraycount = UBound(document)

    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    FState(Me.Tag).Dirty = True
End Sub
Public Sub savefile()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim i As Integer
  
  ' Set CancelError is True
  Dialogs.CancelError = True
  On Error Resume Next
  ' Set flags
  Dialogs.Flags = cdlOFNHideReadOnly
  ' Set filters
  Dialogs.Filter = "All Files (*.*)|*.*|Temperus Files" & _
  "(*.tps)|*.tps"
  ' Specify default filter
  Dialogs.FilterIndex = 2
  ' Display the save dialog box
  Dialogs.ShowSave
  If Err.Number <> 0 Then
    Exit Sub
  End If
  ' get the name file and the path
  name = GetFile(Dialogs.Filename)
  path = GetPath(Dialogs.Filename)

  ' get the current form index
  cur_doc = current_form()
  ' change to the selected directory
  ChDir path

  ReDim material(FState(cur_doc).Conta - 1)
  For i = 1 To FState(cur_doc).Conta - 1
     document(cur_doc).lista.row = i
     document(cur_doc).lista.col = 1
     material(i).l = Val(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 2
     material(i).area = Val(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 3
     material(i).k = Val(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 4
     material(i).q0 = Val(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 5
     material(i).t1 = Val(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 6
     material(i).g0 = Val(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 7
     material(i).n = Val(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 8
     material(i).e = CDbl(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 9
     material(i).alfa = CDbl(document(cur_doc).lista.Text)

  Next i
  
  Open name For Random As #1 Len = Len(material(1))
  With document(cur_doc)
    For i = 1 To FState(cur_doc).Conta - 1
      Put #1, i, material(i)
    Next i
  End With
 Close #1
 FState(cur_doc).saved = True
 document(cur_doc).Caption = name

End Sub


Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
Dim tmp As VbMsgBoxResult

i = current_form()
If FState(i).Conta = 1 Then
    Exit Sub
End If
If Not FState(i).saved Then
  tmp = MsgBox("Save the document: " & Me.Caption & " ?", vbYesNoCancel + vbCritical, " Temperus ")
  If tmp = vbCancel Then
    Cancel = 10
    Exit Sub
  End If
  If tmp = vbYes Then
    Call savefile
    Exit Sub
  End If
End If
FState(i).deleted = True

End Sub

Private Sub form_load()
    lista.MergeCol(2) = True
    lista.MergeCol(5) = True
    lista.MergeCol(6) = True
    lista.ColWidth(0) = TextWidth("###") * 2
    lista.ColWidth(1) = TextWidth("######") * 2
    lista.ColWidth(2) = TextWidth("#######") * 2
    lista.ColWidth(3) = TextWidth("#######") * 2
    lista.ColWidth(4) = TextWidth("#######") * 2
    lista.ColWidth(5) = TextWidth("#####") * 2
    lista.ColWidth(6) = TextWidth("#####") * 2
    lista.ColWidth(7) = TextWidth("#####") * 2
    lista.ColWidth(8) = TextWidth("#####") * 2
    lista.ColWidth(9) = TextWidth("####") * 2
    lista.row = 0
    lista.col = 0
    lista.CellAlignment = 4
    lista.CellFontBold = True
    lista.Text = "nº"
    lista.col = 1
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "L (mm)"
    lista.col = 2
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "Area (mm2)"
    lista.col = 3
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "k (W/m.ºC)"
    lista.col = 4
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "B (W/cm2ºC)"
    lista.col = 5
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "T1 (ºC)"
    lista.col = 6
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "T2 (ºC)"
    lista.col = 7
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "n"
    lista.col = 8
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "E (GPa)"
    lista.col = 9
    lista.CellFontBold = True
    lista.CellFontName = "Symbol"
    lista.CellAlignment = 4
    lista.Text = "a"
    lista.MergeCells = flexMergeRestrictColumns
    SSTab.TabEnabled(1) = False
    SSTab.TabEnabled(2) = False
    SSTab.TabEnabled(3) = False
    SSTab.Tab = 0
    With chart_type
        .AddItem "3dBar"    ' 0
        .AddItem "2dBar"    ' 1
        .AddItem "3dLine"   ' 2
        .AddItem "2dLine"   ' 3
        .AddItem "3dArea"   ' 4
        .AddItem "2dArea"   ' 5
        .AddItem "3dStep"   ' 6
        .AddItem "2dStep"   ' 7
        .AddItem "3dCombination"    ' 8
        .AddItem "2dCombination"    ' 9
        .AddItem "2dPie"    ' 14
        .AddItem "2dXY"     ' 16
        .ListIndex = 11
    End With
    temp_rotation_txt_info.Enabled = False
    temp_rotation_Scroll.Enabled = False
    temp_rotation_txt_angle.Enabled = False
    temp_angle.Enabled = False
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' Show the current form instance as deleted
    FState(Me.Tag).deleted = True
End Sub




Private Sub temp_rotation_Scroll_Change()
    'MSChart1.chartType = VtChChartType3dLine
    temperature_big_chart.Plot.View3d.Rotation = temp_rotation_Scroll.Value
    temp_rotation_txt_angle.Caption = Str(temp_rotation_Scroll.Value) & "º"
    temperature_big_chart.Plot.Axis(VtChAxisIdX).ValueScale.Auto = True

End Sub
