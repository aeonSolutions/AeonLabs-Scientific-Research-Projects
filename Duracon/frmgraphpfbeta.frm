VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmgraphpfbeta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Probability of Failure and Reliability Index Curve"
   ClientHeight    =   4935
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   8295
   Icon            =   "frmgraphpfbeta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Graph Data"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   2175
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Life: "
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
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T ( RI = 1.0 ):"
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
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T ( RI = 1.5 ):"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1080
         TabIndex        =   11
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1320
         TabIndex        =   10
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1320
         TabIndex        =   9
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T ( RI = 3.0 ):"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T ( pf = 10% ): "
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
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   945
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T ( pf = 90% ): "
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
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   945
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1320
         TabIndex        =   5
         Top             =   2160
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1320
         TabIndex        =   4
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "years"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   360
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
      Begin MSChart20Lib.MSChart both_chart 
         Height          =   3765
         Left            =   210
         OleObjectBlob   =   "frmgraphpfbeta.frx":2052
         TabIndex        =   16
         Top             =   390
         Visible         =   0   'False
         Width           =   5355
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit Graph"
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         Top             =   4320
         Width           =   2175
      End
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Begin VB.Menu viewother 
         Caption         =   "&View other graphs"
         Shortcut        =   ^V
      End
      Begin VB.Menu printgraph 
         Caption         =   "&Print graph"
         Shortcut        =   ^P
      End
      Begin VB.Menu CopyGraph 
         Caption         =   "&Copy graph"
         Shortcut        =   ^C
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmgraphpfbeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmgraphpfbeta.Hide
frmgraphchoice.Hide
Unload frmgraphchoice
End Sub

Private Sub CopyGraph_Click()
frmgraphpfbeta.both_chart.EditCopy
End Sub

Private Sub exit_Click()
frmgraphpfbeta.Hide
frmgraphchoice.Hide
Unload frmgraphchoice
End Sub

Private Sub printgraph_Click()
frmgraphpfbeta.PrintForm
End Sub

Private Sub viewother_Click()
frmgraphpfbeta.Hide
frmgraphchoice.Show 1
End Sub
