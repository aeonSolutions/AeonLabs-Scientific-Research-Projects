VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
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
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "temp_rotation_txt_info"
      Tab(1).Control(2)=   "temp_rotation_txt_angle"
      Tab(1).Control(3)=   "temp_angle"
      Tab(1).Control(4)=   "temperature_big_chart"
      Tab(1).Control(5)=   "chart_type"
      Tab(1).Control(6)=   "temp_rotation_Scroll"
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
         Cols            =   11
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
    Dim doc As Integer
    
    doc = current_form
    arraycount = UBound(document)

    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    FState(Me.Tag).Dirty = True
    Me.lista.Refresh
End Sub


Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
Dim tmp As VbMsgBoxResult
Dim j As Integer
Dim name As String
Dim path As String

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
        If FState(i).deleted Then
              MsgBox "ERROR -This message Should not appear!", vbOKCancel, "Info"
            Exit Sub
         End If
         name = FState(i).name
         path = FState(i).path
         If Not FState(i).newname Then
           ' Set CancelError is True
           dialogs.CancelError = True
           On Error Resume Next
           ' Set flags
           dialogs.Flags = cdlOFNHideReadOnly
           ' Set filters
           dialogs.Filter = "All Files (*.*)|*.*|Temperus Files" & _
           "(*.tps)|*.tps"
           ' Specify default filter
           dialogs.FilterIndex = 2
          ' set the working directory the application dir
           dialogs.InitDir = App.path
           ' Display the save dialog box
           dialogs.ShowSave
           If Err.Number <> 0 Then
             Exit Sub
           End If
           ' get the name file and the path
           name = GetFile(dialogs.Filename)
           path = GetPath(dialogs.Filename)
         End If
         Call savefile(name, path, i)
  End If
End If
FState(i).deleted = True
FState(i).Dirty = False
If i > 1 Then
    FState(i - 1).Dirty = True
End If

End Sub

Private Sub form_load()
    ReDim units.b.txt(3)
    ReDim units.b.conversion(3)
          units.b.txt(1) = "W/m2"
          units.b.txt(2) = "W/cm2.ºC"
          units.b.txt(3) = "W/mm2.ºC"
          units.b.conversion(1) = 1
          units.b.conversion(2) = 10000#
          units.b.conversion(3) = 1000000#
          units.b.selected = 1
    
    ReDim units.k.txt(3)
    ReDim units.k.conversion(3)
          units.k.txt(1) = "W/m.ºC"
          units.k.txt(2) = "W/cm.ºC"
          units.k.txt(3) = "W/mm.ºC"
          units.k.conversion(1) = 1
          units.k.conversion(2) = 100#
          units.k.conversion(3) = 1000#
          units.k.selected = 1
    
    ReDim units.te.txt(2)
    ReDim units.te.conversion(2)
          units.te.txt(1) = "ºC"
          units.te.txt(2) = "ºK"
          units.te.conversion(1) = 1
          units.te.conversion(2) = 1
          units.te.selected = 1

    ReDim units.td.txt(2)
    ReDim units.td.conversion(2)
          units.td.txt(1) = "ºC"
          units.td.txt(2) = "ºK"
          units.td.conversion(1) = 1
          units.td.conversion(2) = 1
          units.td.selected = 1
    
    ReDim units.l.txt(5)
    ReDim units.l.conversion(5)
          units.l.txt(1) = "um"
          units.l.txt(2) = "mm"
          units.l.txt(3) = "cm"
          units.l.txt(4) = "dm"
          units.l.txt(5) = "m"
          units.l.conversion(1) = 0.000001
          units.l.conversion(2) = 0.001
          units.l.conversion(3) = 0.01
          units.l.conversion(4) = 0.1
          units.l.conversion(5) = 1
          units.l.selected = 1
    
    ReDim units.q0.txt(4)
    ReDim units.q0.conversion(4)
          units.q0.txt(1) = "W/mm3"
          units.q0.txt(2) = "W/cm3"
          units.q0.txt(3) = "W/dm3"
          units.q0.txt(4) = "W/m3"
          units.q0.conversion(1) = 0.000000001
          units.q0.conversion(2) = 0.000001
          units.q0.conversion(3) = 0.001
          units.q0.conversion(4) = 1
          units.q0.selected = 4
    
    ReDim units.e.txt(2)
    ReDim units.e.conversion(2)
          units.e.txt(1) = "Gpa"
          units.e.txt(2) = "Mpa"
          units.e.conversion(1) = 1000000000#
          units.e.conversion(2) = 1000000#
          units.e.selected = 1
    
    ReDim units.area.txt(5)
    ReDim units.area.conversion(5)
          units.area.txt(1) = "um2"
          units.area.txt(2) = "mm2"
          units.area.txt(3) = "cm2"
          units.area.txt(4) = "dm2"
          units.area.txt(5) = "m2"
          units.area.conversion(1) = 0.000000000001
          units.area.conversion(2) = 0.000001
          units.area.conversion(3) = 0.0001
          units.area.conversion(4) = 0.01
          units.area.conversion(5) = 1
          units.area.selected = 1
  
    lista.MergeCol(2) = True
    lista.MergeCol(5) = True
    lista.MergeCol(6) = True
    lista.ColWidth(0) = TextWidth("###") * 2
    lista.ColWidth(1) = TextWidth("######") * 2
    lista.ColWidth(2) = TextWidth("#######") * 2
    lista.ColWidth(3) = TextWidth("######") * 2
    lista.ColWidth(4) = TextWidth("######") * 2
    lista.ColWidth(5) = TextWidth("####") * 2
    lista.ColWidth(6) = TextWidth("####") * 2
    lista.ColWidth(7) = TextWidth("####") * 2
    lista.ColWidth(8) = TextWidth("####") * 2
    lista.ColWidth(9) = TextWidth("####") * 2
    lista.ColWidth(10) = TextWidth("######") * 2
    lista.row = 0
    lista.col = 0
    lista.CellAlignment = 4
    lista.CellFontBold = True
    lista.Text = "nº"
    lista.col = 1
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "L (" & units.l.txt(units.l.selected) & ")"
    lista.col = 2
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "Area (" & units.area.txt(units.area.selected) & ")"
    lista.col = 3
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "k (" & units.k.txt(units.k.selected) & ")"
    lista.col = 4
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "B (" & units.b.txt(units.b.selected) & ")"
    lista.col = 5
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "Te (" & units.te.txt(units.te.selected) & ")"
    lista.col = 6
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "Td (" & units.td.txt(units.td.selected) & ")"
    lista.col = 7
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "n"
    lista.col = 8
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "E (" & units.e.txt(units.td.selected) & ")"
    lista.col = 9
    lista.CellFontBold = True
    lista.CellFontName = "Symbol"
    lista.CellAlignment = 4
    lista.Text = "a"
    lista.col = 10
    lista.CellFontBold = True
    lista.CellAlignment = 4
    lista.Text = "Q0 (" & units.q0.txt(units.td.selected) & ")"
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
    Dim doc As Integer
    
    doc = current_form
    FState(doc).deleted = True
End Sub



Private Sub temp_rotation_Scroll_Change()
    'MSChart1.chartType = VtChChartType3dLine
    temperature_big_chart.Plot.View3d.Rotation = temp_rotation_Scroll.Value
    temp_rotation_txt_angle.Caption = Str(temp_rotation_Scroll.Value) & "º"
    temperature_big_chart.Plot.Axis(VtChAxisIdX).ValueScale.Auto = True

End Sub
