VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmChild 
   Caption         =   "frmChild"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   Icon            =   "frmChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   540
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dialogs 
      Left            =   300
      Top             =   300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "937095803"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10905
      Left            =   90
      ScaleHeight     =   10905
      ScaleWidth      =   15165
      TabIndex        =   2
      Top             =   180
      Width           =   15165
      Begin VB.Frame Frame6 
         Caption         =   ".: CO2 emissions :."
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
         Left            =   10350
         TabIndex        =   10
         Top             =   6630
         Width           =   4575
         Begin MSChart20Lib.MSChart co2_chart 
            Height          =   3345
            Left            =   30
            OleObjectBlob   =   "frmChild.frx":324A
            TabIndex        =   15
            Top             =   210
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   ".: SO2 emissions :."
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
         Left            =   5340
         TabIndex        =   9
         Top             =   6630
         Width           =   4575
         Begin MSChart20Lib.MSChart so2_chart 
            Height          =   3345
            Left            =   30
            OleObjectBlob   =   "frmChild.frx":4F6E
            TabIndex        =   13
            Top             =   210
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   ".: NOx emissions :."
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
         Left            =   390
         TabIndex        =   8
         Top             =   6630
         Width           =   4575
         Begin MSChart20Lib.MSChart nox_chart 
            Height          =   3345
            Left            =   30
            OleObjectBlob   =   "frmChild.frx":6C92
            TabIndex        =   12
            Top             =   210
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   ".: Water Consuption :."
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
         Left            =   10350
         TabIndex        =   7
         Top             =   2820
         Width           =   4575
         Begin MSChart20Lib.MSChart water_chart 
            Height          =   3345
            Left            =   60
            OleObjectBlob   =   "frmChild.frx":89B6
            TabIndex        =   14
            Top             =   270
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   ".: Energy Consuption :."
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
         Left            =   5340
         TabIndex        =   6
         Top             =   2790
         Width           =   4575
         Begin MSChart20Lib.MSChart energy_chart 
            Height          =   3345
            Left            =   30
            OleObjectBlob   =   "frmChild.frx":A6DA
            TabIndex        =   11
            Top             =   240
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   ".: Structure Costs :."
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
         Left            =   390
         TabIndex        =   4
         Top             =   2790
         Width           =   4575
         Begin MSChart20Lib.MSChart structure_chart 
            Height          =   3345
            Left            =   30
            OleObjectBlob   =   "frmChild.frx":C3FE
            TabIndex        =   5
            Top             =   270
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   2025
         Left            =   450
         TabIndex        =   3
         Top             =   570
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   3572
         _Version        =   393216
         Rows            =   30
         Cols            =   20
         Appearance      =   0
      End
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim HeightDiff As Integer
Dim WidthDiff As Integer

Private Const FORMHEIGHT = 9000
Private Const FORMWIDTH = 10410

Sub Form_Load()
    Picture1.Move 0, 0
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0
    HScroll1.Move 0, Me.ScaleHeight - HScroll1.Height
    VScroll1.Height = Me.ScaleHeight - HScroll1.Height
    HScroll1.Width = Me.ScaleWidth - VScroll1.Width
    '---
    HeightDiff = Picture1.Height - (Me.ScaleHeight - Picture1.top) + HScroll1.Height
    WidthDiff = Picture1.Width - Me.ScaleWidth + VScroll1.Width
    '---
    VScroll1.Min = 1
    VScroll1.Max = HeightDiff
    VScroll1.SmallChange = 100
    VScroll1.LargeChange = 300
    '---
    HScroll1.Min = 1
    HScroll1.Max = WidthDiff
    HScroll1.SmallChange = 100
    HScroll1.LargeChange = 300

    Call load_defaults
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < FORMWIDTH Then Me.Width = FORMWIDTH
    If Err <> 0 Then
        Exit Sub
    End If
    If Me.Height < FORMHEIGHT Then Me.Height = FORMHEIGHT
    'If Me.Width < FORMWIDTH Then Picture1.Width = FORMWIDTH
    'If Me.Height < FORMHEIGHT Then Picture1.Height = FORMHEIGHT
    '---
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0
    HScroll1.Move 0, Me.ScaleHeight - HScroll1.Height
    VScroll1.Height = Me.ScaleHeight - HScroll1.Height
    HScroll1.Width = Me.ScaleWidth - VScroll1.Width
    '---
    If VScroll1.Height >= Picture1.Height Then
        VScroll1.Visible = False
    Else
        VScroll1.Visible = True
    End If
    '---
    If HScroll1.Width >= Picture1.Width Then
        HScroll1.Visible = False
    Else
        HScroll1.Visible = True
    End If
    If VScroll1.Visible = True And HScroll1.Visible = True Then
        VScroll1.Height = Me.ScaleHeight - HScroll1.Height
        HScroll1.Width = Me.ScaleWidth - VScroll1.Width
    ElseIf VScroll1.Visible = False And HScroll1.Visible = True Then
        HScroll1.Width = Me.ScaleWidth
    ElseIf VScroll1.Visible = True And HScroll1.Visible = False Then
        VScroll1.Height = Me.ScaleHeight
    End If
If Me.WindowState = vbMaximized Then
    If Me.Height > Picture1.Height Then
        VScroll1.Visible = False
        VScroll1.Value = 1
    End If
    If Me.Width > Picture1.Width Then
        HScroll1.Visible = False
        HScroll1.Value = 1
    End If
End If
Picture1.SetFocus
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
    If Picture1.Left = 0 Then
        Picture1.Left = HScroll1.Value
    Else
        Picture1.Left = -HScroll1.Value
    End If
    Picture1.SetFocus
End Sub

Private Sub MSChart4_OLEStartDrag(Data As MSChart20Lib.DataObject, AllowedEffects As Long)

End Sub

Private Sub VScroll1_Change()
'=============================
Dim TopMargin As Single

On Error Resume Next

    TopMargin = 0
    If Picture1.top = TopMargin Then
        Picture1.top = Picture1.top - VScroll1.Value
    Else
        Picture1.top = -VScroll1.Value + TopMargin
    End If
    Picture1.SetFocus

End Sub
Sub form_activate()
    Dim arraycount As Integer
    Dim i As Integer
    Dim doc As Integer
    
    doc = Me.Tag
    arraycount = UBound(document)

    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    FState(Me.Tag).Dirty = True
   
End Sub



Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call unload_document
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Show the current form instance as deleted
    Dim doc As Integer
    
    doc = current_form
    FState(doc).deleted = True

End Sub

Private Sub load_defaults()
    Dim doc As Integer
    
    doc = current_form
    
    FState(doc).count = 0
    With Me
        With .lista
            .Row = 0
            .Col = 0
            .ColWidth(0) = TextWidth("#######") * 2
            .Refresh
            .Text = "Type"
            .ColWidth(1) = TextWidth("####") * 2
            .Col = 1
            .CellAlignment = 4
            .Text = "Quantity"
            .ColWidth(2) = TextWidth("####") * 2
            .Col = 2
            .CellAlignment = 4
            .Text = "Height"
            .ColWidth(3) = TextWidth("####") * 2
            .Col = 3
            .CellAlignment = 4
            .Text = "Width"
            .ColWidth(4) = TextWidth("####") * 2
            .Col = 4
            .CellAlignment = 4
            .Text = "Weight"
            .ColWidth(5) = TextWidth("####") * 2
            .Col = 5
            .CellAlignment = 4
            .Text = "Lenght"
            .ColWidth(6) = TextWidth("###") * 2
            .Col = 6
            .CellAlignment = 4
            .Text = "f5"
            .ColWidth(7) = TextWidth("###") * 2
            .Col = 7
            .CellAlignment = 4
            .Text = "f6"
            .ColWidth(8) = TextWidth("###") * 2
            .Col = 8
            .CellAlignment = 4
            .Text = "f8"
            .ColWidth(9) = TextWidth("###") * 2
            .Col = 9
            .CellAlignment = 4
            .Text = "f10"
            .ColWidth(10) = TextWidth("###") * 2
            .Col = 10
            .CellAlignment = 4
            .Text = "f12"
            .ColWidth(11) = TextWidth("###") * 2
            .Col = 11
            .CellAlignment = 4
            .Text = "f16"
            .ColWidth(12) = TextWidth("###") * 2
            .Col = 12
            .CellAlignment = 4
            .Text = "f20"
            .ColWidth(13) = TextWidth("###") * 2
            .Col = 13
            .CellAlignment = 4
            .Text = "f25"
            .ColWidth(14) = TextWidth("###") * 2
            .Col = 14
            .CellAlignment = 4
            .Text = "f32"
            .ColWidth(15) = TextWidth("####") * 2
            .Col = 15
            .CellAlignment = 4
            .Text = "Cement"
            .ColWidth(16) = TextWidth("#####") * 2
            .Col = 16
            .CellAlignment = 4
            .Text = "Aggregates"
            .ColWidth(17) = TextWidth("####") * 2
            .Col = 17
            .CellAlignment = 4
            .Text = "Costs"
            .ColWidth(18) = TextWidth("######") * 2
            .Col = 18
            .CellAlignment = 4
            .Text = "Database"
        End With
    End With

End Sub
