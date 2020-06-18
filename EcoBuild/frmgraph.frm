VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm_graph 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Structure Costs Graph"
   ClientHeight    =   4935
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   8295
   Icon            =   "frmgraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleMode       =   0  'User
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Universidade do Minho"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   6600
      TabIndex        =   1
      Top             =   7320
      Width           =   2895
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Departamento de Engenharia Civil"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sub Grupo de Física das Construções"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2370
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4935
      Index           =   0
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   8265
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin MSChart20Lib.MSChart structure_chart 
         Height          =   3915
         Left            =   1410
         OleObjectBlob   =   "frmgraph.frx":2052
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   5445
      End
   End
   Begin VB.Menu viewother 
      Caption         =   "&View other graphs"
      Begin VB.Menu mnu_global 
         Caption         =   "Global Analysis"
      End
      Begin VB.Menu mnu_energy 
         Caption         =   "Energy Consuption"
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
Attribute VB_Name = "frm_graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim doc As Integer

doc = current_form
Call DisableX(frm_graph(doc))
End Sub


Private Sub Command1_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
End Sub

Private Sub CopyGraph_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).structure_chart.EditCopy
End Sub

Private Sub exit_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
End Sub


Private Sub mnu_co2_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
frm_graph_co2(doc).Show 1


End Sub

Private Sub mnu_energy_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
frm_graph_energy(doc).Show 1

End Sub

Private Sub mnu_global_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
frm_graph_global(doc).Show 1

End Sub

Private Sub mnu_nox_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
frm_graph_nox(doc).Show 1

End Sub

Private Sub mnu_so2_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
frm_graph_so2(doc).Show 1

End Sub

Private Sub mnu_water_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).Hide
frm_graph_water(doc).Show 1

End Sub

Private Sub printgraph_Click()
Dim doc As Integer
doc = current_form

frm_graph(doc).PrintForm
End Sub

