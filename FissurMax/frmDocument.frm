VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   10350
   Icon            =   "frmDocument.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   10350
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   540
      TabIndex        =   58
      Top             =   0
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dialogs 
      Left            =   1320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10575
      Left            =   30
      ScaleHeight     =   10575
      ScaleWidth      =   15165
      TabIndex        =   0
      Top             =   360
      Width           =   15165
      Begin TabDlg.SSTab SSTab 
         Height          =   10095
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   17806
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         TabMaxWidth     =   4410
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Data"
         TabPicture(0)   =   "frmDocument.frx":08CA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "filme"
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(2)=   "num_cracks_txt"
         Tab(0).Control(3)=   "beta_txt"
         Tab(0).Control(4)=   "iteracoes"
         Tab(0).Control(5)=   "weibul"
         Tab(0).Control(6)=   "tensao"
         Tab(0).Control(7)=   "sf_txt"
         Tab(0).Control(8)=   "lowest_txt"
         Tab(0).Control(9)=   "curr_stress_txt"
         Tab(0).Control(10)=   "results"
         Tab(0).Control(11)=   "Label24"
         Tab(0).Control(12)=   "Label19"
         Tab(0).Control(13)=   "Label3"
         Tab(0).Control(14)=   "Label7"
         Tab(0).Control(15)=   "Label25"
         Tab(0).Control(16)=   "Label11"
         Tab(0).Control(17)=   "Label26"
         Tab(0).Control(18)=   "Label27"
         Tab(0).Control(19)=   "Label28"
         Tab(0).ControlCount=   20
         TabCaption(1)   =   "Frequency"
         TabPicture(1)   =   "frmDocument.frx":08E6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "elements_per_segment_txt"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "chart"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Cumulated Frequency"
         TabPicture(2)   =   "frmDocument.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "chart3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Simulated Stress"
         TabPicture(3)   =   "frmDocument.frx":091E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "chart2"
         Tab(3).ControlCount=   1
         Begin VB.Frame filme 
            Caption         =   "Film"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            Left            =   -74820
            TabIndex        =   37
            Top             =   900
            Width           =   2655
            Begin VB.TextBox vf_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   41
               Text            =   "0,3"
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox ef_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   40
               Text            =   "65"
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox l_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   39
               Text            =   "0,125"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox tf_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   38
               Text            =   "0,008"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label20 
               Caption         =   "m"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   46
               Top             =   240
               Width           =   135
            End
            Begin VB.Label Label6 
               Caption         =   "Coef. Poisson"
               Height          =   255
               Left            =   240
               TabIndex        =   45
               Top             =   1710
               Width           =   1035
            End
            Begin VB.Label Label4 
               Caption         =   "E (GPa)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   44
               Top             =   1230
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "L (cm)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   690
               TabIndex        =   43
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label1 
               Caption         =   "Tf (mm)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   540
               TabIndex        =   42
               Top             =   270
               Width           =   735
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Substrate"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   -74820
            TabIndex        =   30
            Top             =   3180
            Width           =   2655
            Begin VB.TextBox vs_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   33
               Text            =   "0,3"
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox es_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   32
               Text            =   "200"
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox ts_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   31
               Text            =   "0,45"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label10 
               Caption         =   "Coef. Poisson"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   90
               TabIndex        =   36
               Top             =   1320
               Width           =   1245
            End
            Begin VB.Label Label9 
               Caption         =   "E (GPa)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   35
               Top             =   870
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "Ts (mm)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   34
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.TextBox num_cracks_txt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -66780
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   510
            Width           =   1215
         End
         Begin VB.TextBox beta_txt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -70080
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   510
            Width           =   1815
         End
         Begin VB.Frame iteracoes 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   -74820
            TabIndex        =   21
            Top             =   8370
            Width           =   2655
            Begin VB.TextBox divisoes_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   24
               Text            =   "3000"
               Top             =   1110
               Width           =   1095
            End
            Begin VB.TextBox n_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   23
               Text            =   "350"
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox segments_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   22
               Text            =   "20"
               Top             =   690
               Width           =   1095
            End
            Begin VB.Label Label18 
               Caption         =   "nº elements"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   210
               TabIndex        =   27
               Top             =   1110
               Width           =   1125
            End
            Begin VB.Label Label17 
               Caption         =   "nº iterations"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   210
               TabIndex        =   26
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label Label5 
               Caption         =   "nº segments"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   210
               TabIndex        =   25
               Top             =   690
               Width           =   1125
            End
         End
         Begin VB.Frame weibul 
            Caption         =   "Weibul"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   -74820
            TabIndex        =   15
            Top             =   6990
            Width           =   2655
            Begin VB.TextBox sigma_weib_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   17
               Text            =   "240"
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox m_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   16
               Text            =   "6"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label21 
               Caption         =   "s"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   9
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   20
               Top             =   840
               Width           =   135
            End
            Begin VB.Label Label16 
               Caption         =   "0 (MPa)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   19
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   "m"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   960
               TabIndex        =   18
               Top             =   360
               Width           =   255
            End
         End
         Begin VB.Frame tensao 
            Caption         =   "Stress"
            Height          =   1815
            Left            =   -74820
            TabIndex        =   6
            Top             =   5100
            Width           =   2655
            Begin VB.TextBox rs_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   9
               Text            =   "0"
               Top             =   1320
               Width           =   1095
            End
            Begin VB.TextBox delta_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   8
               Text            =   "0,5"
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox sigma_txt 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0,0000000000"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2070
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1350
               TabIndex        =   7
               Text            =   "10"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label23 
               Caption         =   "s"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   9
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   14
               Top             =   360
               Width           =   135
            End
            Begin VB.Label Label22 
               Caption         =   "(MPa)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   720
               TabIndex        =   13
               Top             =   840
               Width           =   615
            End
            Begin VB.Label Label14 
               Caption         =   "RS (MPa)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   12
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label13 
               Caption         =   "Ds"
               BeginProperty Font 
                  Name            =   "Symbol"
                  Size            =   8.25
                  Charset         =   2
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   480
               TabIndex        =   11
               Top             =   840
               Width           =   255
            End
            Begin VB.Label Label12 
               Caption         =   "a (MPa)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   10
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.TextBox sf_txt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -64980
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   510
            Width           =   1185
         End
         Begin VB.TextBox lowest_txt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -62010
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   510
            Width           =   1185
         End
         Begin VB.TextBox curr_stress_txt 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0,0000000000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2070
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   -72360
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   540
            Width           =   1185
         End
         Begin MSChart20Lib.MSChart chart 
            Height          =   8865
            Left            =   240
            OleObjectBlob   =   "frmDocument.frx":093A
            TabIndex        =   5
            Top             =   840
            Width           =   14325
         End
         Begin MSFlexGridLib.MSFlexGrid results 
            Height          =   8925
            Left            =   -72060
            TabIndex        =   47
            Top             =   990
            Width           =   11835
            _ExtentX        =   20876
            _ExtentY        =   15743
            _Version        =   393216
            Rows            =   3000
            Cols            =   12
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
            MergeCells      =   3
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
         Begin MSChart20Lib.MSChart chart3 
            Height          =   9225
            Left            =   -74880
            OleObjectBlob   =   "frmDocument.frx":215C
            TabIndex        =   60
            Top             =   720
            Width           =   14325
         End
         Begin MSChart20Lib.MSChart chart2 
            Height          =   9225
            Left            =   -74640
            OleObjectBlob   =   "frmDocument.frx":44B4
            TabIndex        =   61
            Top             =   600
            Width           =   14325
         End
         Begin VB.Label Label24 
            Caption         =   "nº of cracks :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -68010
            TabIndex        =   57
            Top             =   540
            Width           =   1245
         End
         Begin VB.Label Label19 
            Caption         =   "b :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -70320
            TabIndex        =   56
            Top             =   540
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "s   :"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   9.75
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -65370
            TabIndex        =   55
            Top             =   510
            Width           =   345
         End
         Begin VB.Label Label7 
            Caption         =   "F"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   -65250
            TabIndex        =   54
            Top             =   660
            Width           =   135
         End
         Begin VB.Label Label25 
            Caption         =   "MPa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -63750
            TabIndex        =   53
            Top             =   540
            Width           =   495
         End
         Begin VB.Label elements_per_segment_txt 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   720
            TabIndex        =   52
            Top             =   480
            Width           =   4365
         End
         Begin VB.Label Label11 
            Caption         =   "Lowest Random"
            Height          =   255
            Left            =   -63240
            TabIndex        =   51
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "MPa"
            Height          =   225
            Left            =   -60780
            TabIndex        =   50
            Top             =   540
            Width           =   405
         End
         Begin VB.Label Label27 
            Caption         =   "External Stress:"
            Height          =   285
            Left            =   -73530
            TabIndex        =   49
            Top             =   570
            Width           =   1155
         End
         Begin VB.Label Label28 
            Caption         =   "MPa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -71100
            TabIndex        =   48
            Top             =   570
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim HeightDiff As Integer
Dim WidthDiff As Integer

Private Const FORMHEIGHT = 10095
Private Const FORMWIDTH = 14910

Private Sub form_load()

    Call colwidth
    SSTab.TabEnabled(1) = False
    SSTab.TabEnabled(2) = False
    SSTab.TabEnabled(3) = False
    SSTab.Tab = 0
    generated = False
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
           Dialogs.CancelError = True
           On Error Resume Next
           ' Set flags
           Dialogs.Flags = cdlOFNHideReadOnly
           ' Set filters
           Dialogs.Filter = "All Files (*.*)|*.*|FissurMax Files" & _
           "(*.fmx)|*.fmx"
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

Private Sub Form_Unload(Cancel As Integer)
    ' Show the current form instance as deleted
    Dim doc As Integer
    
    doc = current_form
    FState(doc).deleted = True
End Sub

Private Sub MDIForm_GotFocus()
    Call colwidth
    Call load_results
End Sub

Private Sub l_txt_Change()
    generated = False
    FState(Me.Tag).saved = False
End Sub

Private Sub delta_txt_Change()
    FState(Me.Tag).saved = False
End Sub

Private Sub divisoes_txt_Change()
    FState(Me.Tag).saved = False
    generated = False
End Sub

Private Sub ef_txt_Change()
    FState(Me.Tag).saved = False
End Sub

Private Sub es_txt_Change()
    FState(Me.Tag).saved = False
End Sub

Private Sub m_txt_Change()
    generated = False
    FState(Me.Tag).saved = False
End Sub

Private Sub n_txt_Change()
    FState(Me.Tag).saved = False
End Sub

Private Sub rs_txt_Change()
    FState(Me.Tag).saved = False
End Sub

Private Sub segments_txt_Change()
    FState(Me.Tag).saved = False
End Sub

Private Sub sigma_txt_Change()
    generated = False
    FState(Me.Tag).saved = False
End Sub

Private Sub sigma_weib_txt_Change()
    generated = False
    FState(Me.Tag).saved = False
End Sub

Private Sub tf_txt_Change()

    FState(Me.Tag).saved = False
End Sub

Private Sub ts_txt_Change()

    FState(Me.Tag).saved = False
End Sub

Private Sub vf_txt_Change()

    FState(Me.Tag).saved = False
End Sub

Private Sub vs_txt_Change()

    FState(Me.Tag).saved = False
End Sub
Sub colwidth()
    results.Row = 0
    results.Col = 0
    results.MergeCol(5) = True
    results.MergeCol(6) = True
    results.MergeCol(7) = True
    results.MergeCol(8) = True
    results.Refresh
    results.Text = "nº"
    results.colwidth(1) = TextWidth("#########") * 2
    results.Col = 1
    results.CellAlignment = 4
    results.Text = "current stress(MPa)"
    results.colwidth(2) = TextWidth("########") * 2
    results.Col = 2
    results.CellAlignment = 4
    results.Text = "Final stress(MPa)"
    results.colwidth(3) = TextWidth("######") * 2
    results.Col = 3
    results.CellAlignment = 4
    results.Text = "Cracked ?"
    results.colwidth(4) = TextWidth("######") * 2
    results.Col = 4
    results.CellAlignment = 4
    results.Text = "Coord (mm)"
    results.colwidth(5) = TextWidth("######") * 2
    results.Col = 5
    results.CellAlignment = 4
    results.Text = "L (mm)"
    results.colwidth(6) = TextWidth("######") * 2
    results.Col = 6
    results.CellAlignment = 4
    results.Text = "initial pos"
    results.colwidth(7) = TextWidth("#####") * 2
    results.Col = 7
    results.CellAlignment = 4
    results.Text = "Final pos"
    results.colwidth(8) = TextWidth("#####") * 2
    results.Col = 8
    results.CellAlignment = 4
    results.Text = "nº blocks"
End Sub
