VERSION 5.00
Begin VB.Form Print_graph 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PRINT PREVIEW"
   ClientHeight    =   6270
   ClientLeft      =   2295
   ClientTop       =   3060
   ClientWidth     =   8970
   Icon            =   "print_graph.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   14
      TabIndex        =   4
      Top             =   5160
      Width           =   9015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Width           =   9015
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "P&rint"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Print"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel"
         Top             =   120
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Chart Dim"
         Height          =   855
         Left            =   3600
         TabIndex        =   9
         Top             =   120
         Width           =   3495
         Begin VB.HScrollBar chtscroll 
            Height          =   255
            Index           =   1
            LargeChange     =   5
            Left            =   1800
            Max             =   100
            Min             =   25
            TabIndex        =   11
            Top             =   480
            Value           =   25
            Width           =   1575
         End
         Begin VB.HScrollBar chtscroll 
            Height          =   255
            Index           =   0
            LargeChange     =   5
            Left            =   240
            Max             =   100
            Min             =   25
            TabIndex        =   10
            Top             =   480
            Value           =   25
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Width"
            Height          =   255
            Left            =   1800
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Height"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ZOOM"
         Height          =   855
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   3615
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            LargeChange     =   5
            Left            =   120
            Max             =   100
            Min             =   25
            TabIndex        =   7
            Top             =   360
            Value           =   25
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "%"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   3120
            TabIndex        =   8
            Top             =   360
            Width           =   375
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5175
      Left            =   8760
      Max             =   14
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   17000
      Left            =   0
      ScaleHeight     =   16965
      ScaleWidth      =   11970
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         MousePointer    =   2  'Cross
         ScaleHeight     =   1065
         ScaleWidth      =   2265
         TabIndex        =   1
         ToolTipText     =   "Drag and drop anywhere on the page"
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   0
      X2              =   9000
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu op1 
      Caption         =   "Options"
      Index           =   0
      Begin VB.Menu opt1 
         Caption         =   "Paper Size"
         Index           =   0
         Begin VB.Menu A4 
            Caption         =   "A4"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu or 
         Caption         =   "Orientation"
         Begin VB.Menu or1 
            Caption         =   "Portrait/Landscape"
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu pq 
         Caption         =   "Print Quality"
         Begin VB.Menu pq1 
            Caption         =   "Low"
            Index           =   0
         End
         Begin VB.Menu pq1 
            Caption         =   "Medium"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu pq1 
            Caption         =   "High"
            Index           =   2
         End
      End
      Begin VB.Menu noc 
         Caption         =   "No Of Copies"
         Begin VB.Menu noc1 
            Caption         =   "1"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu noc1 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu noc1 
            Caption         =   "3"
            Index           =   2
         End
         Begin VB.Menu noc1 
            Caption         =   "4"
            Index           =   3
         End
         Begin VB.Menu noc1 
            Caption         =   "5"
            Index           =   4
         End
         Begin VB.Menu noc1 
            Caption         =   "6"
            Index           =   5
         End
         Begin VB.Menu noc1 
            Caption         =   "7"
            Index           =   6
         End
         Begin VB.Menu noc1 
            Caption         =   "8"
            Index           =   7
         End
         Begin VB.Menu noc1 
            Caption         =   "9"
            Index           =   8
         End
         Begin VB.Menu noc1 
            Caption         =   "10"
            Index           =   9
         End
      End
   End
End
Attribute VB_Name = "Print_graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim px As Long
Dim py As Long
Dim mos As Boolean
Dim h As Long
Option Explicit

Dim w As Long
Dim l1 As Long
Dim t1 As Long
Dim perc As Double
Dim orientint As Long
Dim prqltint As Long
Dim nocopyint As Integer
Dim hsc2 As Long
Dim doc As Integer




Private Sub chtscroll_Change(Index As Integer)
Dim dd As Printer
  
doc = current_form

  Select Case Index
    Case 0
      frm_exp_data(doc).chart_exp_data.Height = ((h * (chtscroll(Index).Value / 100)) / 0.25)
      frm_exp_data(doc).chht = (frm_exp_data(doc).chart_exp_data.Height / (HScroll2.Value / 100))
    Case 1
      frm_exp_data(doc).chart_exp_data.Width = ((w * (chtscroll(Index).Value / 100)) / 0.25)
      frm_exp_data(doc).chwd = (frm_exp_data(doc).chart_exp_data.Width / (HScroll2.Value / 100))
  End Select
  Clipboard.Clear
  frm_exp_data(doc).chart_exp_data.EditCopy
  Print_graph.Picture2.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub cmdcancel_Click()
  Clipboard.Clear
  Unload Me
End Sub

Private Sub cmdPrint_Click()
  Dim schval As Long
  Dim obj As Printer
  Dim hh As Variant
  Dim tt As Variant
  
  schval = HScroll2.Value
  HScroll2.Value = HScroll2.Max
  hh = Picture2.Top
  tt = Picture2.Left
  Call printchart
End Sub
Private Sub printchart()
  Dim i As Integer
  
  With Printer
  
    .PaperSize = vbPRPSA4
    For i = 1 To nocopyint
        .PaintPicture Clipboard.GetData(vbCFDIB), Picture2.Left, Picture2.Top
        .NewPage
    Next
    .EndDoc
End With
End Sub

Private Sub Form_Load()
  Call Centerform(Print_graph)
  Print_graph.Picture2.Picture = Clipboard.GetData(vbCFDIB)
  
  h = Picture2.Height
  w = Picture2.Width
  Picture1.Left = 0
  Picture1.Top = 0
  'Initialising The Recordset
'  Set mainheadrs = New Recordset
'  mainheadrs.Open "select * from temp", dbname1, adOpenStatic, adLockOptimistic
                   
  
  HScroll2.Value = 25
  perc = 0.25
  HScroll2_Change

Call DisableX(Print_graph)

End Sub

Private Sub HScroll2_Change()

  Label1.Caption = HScroll2.Value & "%"
  If perc = 0 Then Exit Sub
  
  If or1(0).Checked = True Then
    Picture1.Width = 12000 * (HScroll2.Value / 100)
    Picture1.Height = 17000 * (HScroll2.Value / 100)
  Else
    Picture1.Width = 17000 * (HScroll2.Value / 100)
    Picture1.Height = 12000 * (HScroll2.Value / 100)
  End If
  
  doc = current_form

  frm_exp_data(doc).chart_exp_data.Width = frm_exp_data(doc).chwd * (HScroll2.Value / 100)
  frm_exp_data(doc).chart_exp_data.Height = frm_exp_data(doc).chht * (HScroll2.Value / 100)
  h = frm_exp_data(doc).chart_exp_data.Height
  w = frm_exp_data(doc).chart_exp_data.Width
  Clipboard.Clear
  frm_exp_data(doc).chart_exp_data.EditCopy
  Print_graph.Picture2.Picture = Clipboard.GetData(vbCFDIB)
  Picture2.Left = ((l1 * (HScroll2.Value / 100)) / perc)
  Picture2.Top = (t1 * (HScroll2.Value / 100)) / perc

End Sub

Private Sub noc1_Click(Index As Integer)
Dim i As Integer

  nocopyint = Index + 1
  For i = 0 To Me.noc1.Count - 1
    Me.noc1(i).Checked = False
  Next
  noc1(Index).Checked = True
End Sub

Private Sub or1_Click(Index As Integer)
  Dim wx As Long
  Dim hy As Long
  wx = Picture1.Width
  hy = Picture1.Height
  Select Case Index
  Case 0
    orientint = 2
    Picture1.Width = hy
    Picture1.Height = wx
    If or1(Index).Checked = True Then
      or1(Index).Checked = False
    Else
      or1(Index).Checked = True
    End If
  End Select
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  mos = True
  px = x
  py = Y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If mos = True Then
    Picture2.Left = Picture2.Left + x - px
    Picture2.Top = Picture2.Top + Y - py
  End If
End Sub

Private Sub pq1_Click(Index As Integer)
  Dim mn As Menu
  Dim i As Integer
  
  For i = 0 To Me.pq1.Count - 1
    Me.pq1(i).Checked = False
  Next
  Me.pq1(Index).Checked = True
  Select Case Index
  Case 0
    prqltint = -2
  Case 1
    prqltint = -3
  Case 2
    prqltint = -4
  End Select
End Sub

Private Sub VScroll1_Change()
  Picture1.Top = 0 - VScroll1.Value * 900
End Sub
Private Sub HScroll1_Change()
  Picture1.Left = 0 - HScroll1.Value * 400
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  mos = False
  Picture2.Left = Picture2.Left + x - px
  Picture2.Top = Picture2.Top + Y - py
  l1 = Picture2.Left
  t1 = Picture2.Top
  perc = HScroll2.Value / 100
End Sub
