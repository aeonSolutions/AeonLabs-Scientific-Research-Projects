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
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   540
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dialogs 
      Left            =   60
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "937095803"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10785
      Left            =   90
      ScaleHeight     =   10785
      ScaleWidth      =   15165
      TabIndex        =   2
      Top             =   0
      Width           =   15165
      Begin VB.Frame Frame4 
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
         Left            =   5040
         TabIndex        =   17
         Top             =   6960
         Width           =   2175
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( RI = 4.0 ): "
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
            TabIndex        =   30
            Top             =   2160
            Width           =   900
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( RI = 3.0 ): "
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
            TabIndex        =   29
            Top             =   1800
            Width           =   900
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( RI = 2.0 ): "
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
            TabIndex        =   28
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( RI = 1.5 ): "
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
            TabIndex        =   27
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( RI = 1.0 ): "
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
            TabIndex        =   26
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label1 
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
            TabIndex        =   25
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label4 
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
            TabIndex        =   24
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   23
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   22
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   21
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   20
            Top             =   1800
            Width           =   135
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   19
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label Label26 
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
            TabIndex        =   18
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   ".: Probability of failure vs time :."
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
         Left            =   7620
         TabIndex        =   15
         Top             =   6390
         Width           =   4575
         Begin MSChart20Lib.MSChart pf_chart 
            Height          =   3345
            Left            =   60
            OleObjectBlob   =   "frmChild.frx":324A
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   4485
         End
      End
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
         Left            =   12600
         TabIndex        =   6
         Top             =   6960
         Width           =   2175
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( pf = 99% ): "
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
            TabIndex        =   35
            Top             =   2160
            Width           =   945
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( pf = 95% ): "
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
            TabIndex        =   34
            Top             =   1800
            Width           =   945
         End
         Begin VB.Label Label20 
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
            TabIndex        =   33
            Top             =   1440
            Width           =   945
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T ( pf = 50% ): "
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
            TabIndex        =   32
            Top             =   1080
            Width           =   945
         End
         Begin VB.Label Label2 
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
            TabIndex        =   31
            Top             =   720
            Width           =   945
         End
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
            TabIndex        =   13
            Top             =   360
            Width           =   225
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   12
            Top             =   1080
            Width           =   135
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   11
            Top             =   720
            Width           =   135
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            Top             =   1800
            Width           =   135
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " - "
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
            TabIndex        =   8
            Top             =   1440
            Width           =   135
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
            TabIndex        =   7
            Top             =   360
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   ".: Reliability index vs Time :."
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
         TabIndex        =   4
         Top             =   6420
         Width           =   4575
         Begin MSChart20Lib.MSChart reliability_chart 
            Height          =   3345
            Left            =   30
            OleObjectBlob   =   "frmChild.frx":577D
            TabIndex        =   5
            Top             =   270
            Visible         =   0   'False
            Width           =   4485
         End
      End
      Begin MSFlexGridLib.MSFlexGrid lista 
         Height          =   6105
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   10769
         _Version        =   393216
         Rows            =   60
         Cols            =   50
         FixedRows       =   0
         FixedCols       =   0
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         GridLines       =   0
         GridLinesFixed  =   0
         MergeCells      =   1
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

Private Const FORMHEIGHT = 9510
Private Const FORMWIDTH = 10530

Sub Form_Load()
    Picture1.Move 0, 0
    VScroll1.Move Me.ScaleWidth - VScroll1.Width, 0
    HScroll1.Move 0, Me.ScaleHeight - HScroll1.Height
    VScroll1.Height = Me.ScaleHeight - HScroll1.Height
    HScroll1.Width = Me.ScaleWidth - VScroll1.Width
    '---
    HeightDiff = Picture1.Height - (Me.ScaleHeight - Picture1.Top) + HScroll1.Height
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

End Sub
Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < FORMWIDTH Then Me.Width = FORMWIDTH
    If Err <> 0 Then
        Exit Sub
    End If
    If Me.Height < FORMHEIGHT Then Me.Height = FORMHEIGHT
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
    End If
    If Me.Width > Picture1.Width Then
        HScroll1.Visible = False
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

Private Sub VScroll1_Change()
'=============================
Dim TopMargin As Single

On Error Resume Next

    TopMargin = 0
    If Picture1.Top = TopMargin Then
        Picture1.Top = Picture1.Top - VScroll1.Value
    Else
        Picture1.Top = -VScroll1.Value + TopMargin
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
    Me.lista.MergeCells = flexMergeFree

    Call refresh_lista(doc)
    
    Me.lista.Refresh
    
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

