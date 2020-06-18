VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_database 
   Caption         =   "Database Maintenance"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   Icon            =   "Frm_database.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   570
      TabIndex        =   48
      Top             =   0
      Width           =   1335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin MSComDlg.CommonDialog dialogs 
      Left            =   330
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
      Left            =   60
      ScaleHeight     =   10905
      ScaleWidth      =   15165
      TabIndex        =   49
      Top             =   30
      Width           =   15165
      Begin TabDlg.SSTab SSTab 
         Height          =   10335
         Left            =   90
         TabIndex        =   50
         Top             =   60
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   18230
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Concrete"
         TabPicture(0)   =   "Frm_database.frx":324A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame4"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame5"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame9"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Metalic"
         TabPicture(1)   =   "Frm_database.frx":3266
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(1)=   "Frame7"
         Tab(1).Control(2)=   "Frame8"
         Tab(1).Control(3)=   "Timer1"
         Tab(1).ControlCount=   4
         Begin VB.Frame Frame9 
            Caption         =   "Water Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   4560
            TabIndex        =   131
            Top             =   3360
            Width           =   5865
            Begin VB.TextBox water_ec_text 
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
               Left            =   2400
               TabIndex        =   8
               Text            =   "1"
               Top             =   330
               Width           =   1545
            End
            Begin VB.TextBox water_water_text 
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
               Left            =   2400
               TabIndex        =   9
               Text            =   "1"
               Top             =   690
               Width           =   1545
            End
            Begin VB.TextBox water_co2_text 
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
               Left            =   2400
               TabIndex        =   132
               Text            =   "1"
               Top             =   1050
               Width           =   1545
            End
            Begin VB.TextBox water_so2_text 
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
               Left            =   2400
               TabIndex        =   10
               Text            =   "1"
               Top             =   1410
               Width           =   1545
            End
            Begin VB.TextBox water_nox_text 
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
               Left            =   2400
               TabIndex        =   11
               Text            =   "1"
               Top             =   1770
               Width           =   1545
            End
            Begin VB.Label Label80 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy Consuption"
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
               Left            =   450
               TabIndex        =   142
               Top             =   360
               Width           =   1875
            End
            Begin VB.Label Label79 
               Alignment       =   1  'Right Justify
               Caption         =   "Water"
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
               Left            =   1560
               TabIndex        =   141
               Top             =   690
               Width           =   765
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               Caption         =   "CO2"
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
               Left            =   1350
               TabIndex        =   140
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
               Caption         =   "SO2"
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
               Left            =   1380
               TabIndex        =   139
               Top             =   1440
               Width           =   945
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
               Caption         =   "NOx"
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
               Left            =   1200
               TabIndex        =   138
               Top             =   1800
               Width           =   1155
            End
            Begin VB.Label Label75 
               Caption         =   "GJ/ton"
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
               Left            =   4050
               TabIndex        =   137
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label64 
               Caption         =   "m3/ton"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   4050
               TabIndex        =   136
               Top             =   720
               Width           =   885
            End
            Begin VB.Label Label59 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   135
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label51 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   134
               Top             =   1440
               Width           =   915
            End
            Begin VB.Label Label44 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   133
               Top             =   1800
               Width           =   705
            End
         End
         Begin VB.Timer Timer1 
            Interval        =   500
            Left            =   -74280
            Top             =   240
         End
         Begin VB.Frame Frame8 
            Caption         =   "Steel Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Left            =   -73830
            TabIndex        =   120
            Top             =   6480
            Width           =   5865
            Begin VB.TextBox metalic_ec_text 
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
               Left            =   2400
               TabIndex        =   39
               Top             =   480
               Width           =   1545
            End
            Begin VB.TextBox metalic_water_text 
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
               Left            =   2400
               TabIndex        =   40
               Top             =   810
               Width           =   1545
            End
            Begin VB.TextBox metalic_co2_text 
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
               Left            =   2400
               TabIndex        =   41
               Top             =   1170
               Width           =   1545
            End
            Begin VB.TextBox metalic_so2_text 
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
               Left            =   2400
               TabIndex        =   42
               Top             =   1530
               Width           =   1545
            End
            Begin VB.TextBox metalic_nox_text 
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
               Left            =   2400
               TabIndex        =   43
               Top             =   1890
               Width           =   1545
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy Consuption"
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
               Left            =   450
               TabIndex        =   130
               Top             =   480
               Width           =   1875
            End
            Begin VB.Label Label73 
               Alignment       =   1  'Right Justify
               Caption         =   "Water"
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
               Left            =   1560
               TabIndex        =   129
               Top             =   810
               Width           =   765
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               Caption         =   "CO2"
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
               Left            =   1350
               TabIndex        =   128
               Top             =   1200
               Width           =   975
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
               Caption         =   "SO2"
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
               Left            =   1380
               TabIndex        =   127
               Top             =   1560
               Width           =   945
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               Caption         =   "NOx"
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
               Left            =   1200
               TabIndex        =   126
               Top             =   1920
               Width           =   1155
            End
            Begin VB.Label Label69 
               Caption         =   "GJ/ton"
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
               Left            =   4050
               TabIndex        =   125
               Top             =   480
               Width           =   885
            End
            Begin VB.Label Label68 
               Caption         =   "m3/ton"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   4050
               TabIndex        =   124
               Top             =   840
               Width           =   885
            End
            Begin VB.Label Label67 
               Caption         =   "ton/ton"
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
               Left            =   4050
               TabIndex        =   123
               Top             =   1200
               Width           =   1035
            End
            Begin VB.Label Label66 
               Caption         =   "ton/ton"
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
               Left            =   4050
               TabIndex        =   122
               Top             =   1560
               Width           =   915
            End
            Begin VB.Label Label65 
               Caption         =   "ton/ton"
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
               Left            =   4050
               TabIndex        =   121
               Top             =   1920
               Width           =   705
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Transportation Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2475
            Left            =   -66930
            TabIndex        =   111
            Top             =   6480
            Width           =   5865
            Begin VB.TextBox metalic_distance_text 
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
               Left            =   2400
               TabIndex        =   44
               Top             =   570
               Width           =   1545
            End
            Begin VB.TextBox trans_co2_text 
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
               Left            =   2400
               TabIndex        =   45
               Top             =   930
               Width           =   1545
            End
            Begin VB.TextBox trans_so2_text 
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
               Left            =   2400
               TabIndex        =   46
               Top             =   1290
               Width           =   1545
            End
            Begin VB.TextBox trans_nox_text 
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
               Left            =   2400
               TabIndex        =   47
               Top             =   1650
               Width           =   1545
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               Caption         =   "Distance"
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
               Left            =   1560
               TabIndex        =   119
               Top             =   600
               Width           =   765
            End
            Begin VB.Label Label62 
               Alignment       =   1  'Right Justify
               Caption         =   "CO2"
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
               Left            =   1350
               TabIndex        =   118
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               Caption         =   "SO2"
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
               Left            =   1380
               TabIndex        =   117
               Top             =   1320
               Width           =   945
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               Caption         =   "NOx"
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
               Left            =   1170
               TabIndex        =   116
               Top             =   1650
               Width           =   1155
            End
            Begin VB.Label Label58 
               Caption         =   "Km"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   4050
               TabIndex        =   115
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label57 
               Caption         =   "Gton/Km"
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
               Left            =   4050
               TabIndex        =   114
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label Label56 
               Caption         =   "Gton/Km"
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
               Left            =   4050
               TabIndex        =   113
               Top             =   1320
               Width           =   915
            End
            Begin VB.Label Label55 
               Caption         =   "Gton/Km"
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
               Left            =   4050
               TabIndex        =   112
               Top             =   1680
               Width           =   885
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "General Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4185
            Left            =   -73830
            TabIndex        =   103
            Top             =   1650
            Width           =   12765
            Begin VB.CommandButton metal_copy_button 
               Caption         =   "Copy to New Entry"
               Enabled         =   0   'False
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
               Left            =   10320
               TabIndex        =   35
               Top             =   1020
               Width           =   2025
            End
            Begin VB.TextBox metalic_name_text 
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
               Left            =   2340
               MaxLength       =   20
               TabIndex        =   36
               Top             =   1830
               Width           =   2865
            End
            Begin VB.TextBox metalic_date_text 
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
               Left            =   2340
               TabIndex        =   37
               Top             =   2220
               Width           =   1545
            End
            Begin VB.TextBox metalic_description_text 
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
               Left            =   2340
               MaxLength       =   255
               ScrollBars      =   2  'Vertical
               TabIndex        =   38
               Top             =   2610
               Width           =   10215
            End
            Begin VB.ComboBox metalic_entries 
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
               Height          =   315
               Left            =   4950
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   990
               Width           =   1905
            End
            Begin VB.CommandButton metalic_save_button 
               Caption         =   "Save"
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
               Left            =   7320
               TabIndex        =   33
               Top             =   1020
               Width           =   1215
            End
            Begin VB.CommandButton metalic_delete_button 
               Caption         =   "Delete"
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
               Left            =   8820
               TabIndex        =   34
               Top             =   1020
               Width           =   1215
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
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
               Left            =   1290
               TabIndex        =   110
               Top             =   1860
               Width           =   975
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               Caption         =   "Date"
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
               Left            =   1500
               TabIndex        =   109
               Top             =   2250
               Width           =   765
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               Caption         =   "Description"
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
               Left            =   870
               TabIndex        =   108
               Top             =   2640
               Width           =   1395
            End
            Begin VB.Label num_metalic_entries 
               Caption         =   "There are [CODED] entries in the Metalic database"
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
               Left            =   570
               TabIndex        =   107
               Top             =   420
               Width           =   6795
            End
            Begin VB.Label Label50 
               Caption         =   "(20 characters max)"
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
               Left            =   5340
               TabIndex        =   106
               Top             =   1890
               Width           =   2355
            End
            Begin VB.Label Label49 
               Caption         =   "Please select  the entry you wish to modify / create in the dropdown box"
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
               Left            =   570
               TabIndex        =   105
               Top             =   630
               Width           =   6795
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   "Selected Entry"
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
               Left            =   2970
               TabIndex        =   104
               Top             =   1020
               Width           =   1875
            End
            Begin VB.Line Line2 
               BorderColor     =   &H8000000C&
               X1              =   330
               X2              =   12390
               Y1              =   1530
               Y2              =   1530
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "General Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2805
            Left            =   1170
            TabIndex        =   95
            Top             =   420
            Width           =   12765
            Begin VB.CommandButton concrete_copy_button 
               Caption         =   "Copy to New Entry"
               Enabled         =   0   'False
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
               Left            =   10290
               TabIndex        =   4
               Top             =   1020
               Width           =   2025
            End
            Begin VB.CommandButton concrete_delete_button 
               Caption         =   "Delete"
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
               Left            =   8820
               TabIndex        =   3
               Top             =   1020
               Width           =   1215
            End
            Begin VB.CommandButton concrete_save_button 
               Caption         =   "Save"
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
               Left            =   7290
               TabIndex        =   2
               Top             =   1020
               Width           =   1215
            End
            Begin VB.ComboBox concrete_entries 
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
               Height          =   315
               Left            =   4950
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   990
               Width           =   1905
            End
            Begin VB.TextBox concrete_description_text 
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
               Left            =   2340
               MaxLength       =   255
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Text            =   "test"
               Top             =   2370
               Width           =   10215
            End
            Begin VB.TextBox concrete_date_text 
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
               Left            =   2340
               MaxLength       =   10
               TabIndex        =   6
               Top             =   1980
               Width           =   1545
            End
            Begin VB.TextBox concrete_name_text 
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
               Left            =   2340
               MaxLength       =   20
               TabIndex        =   5
               Text            =   "teste"
               Top             =   1590
               Width           =   2865
            End
            Begin VB.Line Line1 
               BorderColor     =   &H8000000C&
               X1              =   330
               X2              =   12390
               Y1              =   1410
               Y2              =   1410
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               Caption         =   "Selected Entry"
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
               Left            =   2970
               TabIndex        =   102
               Top             =   1020
               Width           =   1875
            End
            Begin VB.Label Label46 
               Caption         =   "Please select  the entry you wish to modify / create in the dropdown box"
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
               Left            =   570
               TabIndex        =   101
               Top             =   630
               Width           =   6795
            End
            Begin VB.Label Label45 
               Caption         =   "(20 characters max)"
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
               Left            =   5340
               TabIndex        =   100
               Top             =   1650
               Width           =   2355
            End
            Begin VB.Label num_concrete_entries 
               Caption         =   "There are [CODED] entries in the concrete database"
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
               Left            =   570
               TabIndex        =   99
               Top             =   420
               Width           =   6795
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               Caption         =   "Description"
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
               Left            =   870
               TabIndex        =   98
               Top             =   2400
               Width           =   1395
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               Caption         =   "Date"
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
               Left            =   1500
               TabIndex        =   97
               Top             =   2010
               Width           =   765
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               Caption         =   "Name"
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
               Left            =   1290
               TabIndex        =   96
               Top             =   1620
               Width           =   975
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Steel Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   8040
            TabIndex        =   84
            Top             =   7920
            Width           =   5865
            Begin VB.TextBox steel_nox_text 
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
               Left            =   2400
               TabIndex        =   31
               Text            =   "1"
               Top             =   1770
               Width           =   1545
            End
            Begin VB.TextBox steel_so2_text 
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
               Left            =   2400
               TabIndex        =   30
               Text            =   "1"
               Top             =   1410
               Width           =   1545
            End
            Begin VB.TextBox steel_co2_text 
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
               Left            =   2400
               TabIndex        =   29
               Text            =   "1"
               Top             =   1050
               Width           =   1545
            End
            Begin VB.TextBox steel_water_text 
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
               Left            =   2400
               TabIndex        =   28
               Text            =   "1"
               Top             =   690
               Width           =   1545
            End
            Begin VB.TextBox steel_ec_text 
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
               Left            =   2400
               TabIndex        =   27
               Text            =   "1"
               Top             =   330
               Width           =   1545
            End
            Begin VB.Label Label40 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   94
               Top             =   1800
               Width           =   705
            End
            Begin VB.Label Label39 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   93
               Top             =   1440
               Width           =   915
            End
            Begin VB.Label Label38 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   92
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label37 
               Caption         =   "m3/ton"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   4050
               TabIndex        =   91
               Top             =   720
               Width           =   885
            End
            Begin VB.Label Label36 
               Caption         =   "GJ/ton"
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
               Left            =   4050
               TabIndex        =   90
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "NOx"
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
               Left            =   1200
               TabIndex        =   89
               Top             =   1800
               Width           =   1155
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               Caption         =   "SO2"
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
               Left            =   1380
               TabIndex        =   88
               Top             =   1440
               Width           =   945
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               Caption         =   "CO2"
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
               Left            =   1350
               TabIndex        =   87
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               Caption         =   "Water"
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
               Left            =   1560
               TabIndex        =   86
               Top             =   690
               Width           =   765
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy Consuption"
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
               Left            =   450
               TabIndex        =   85
               Top             =   360
               Width           =   1875
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Wood Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   8070
            TabIndex        =   73
            Top             =   5640
            Width           =   5865
            Begin VB.TextBox wood_nox_text 
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
               Left            =   2400
               TabIndex        =   26
               Text            =   "1"
               Top             =   1770
               Width           =   1545
            End
            Begin VB.TextBox wood_so2_text 
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
               Left            =   2400
               TabIndex        =   25
               Text            =   "1"
               Top             =   1410
               Width           =   1545
            End
            Begin VB.TextBox wood_co2_text 
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
               Left            =   2400
               TabIndex        =   24
               Text            =   "1"
               Top             =   1050
               Width           =   1545
            End
            Begin VB.TextBox wood_water_text 
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
               Left            =   2400
               TabIndex        =   23
               Text            =   "1"
               Top             =   690
               Width           =   1545
            End
            Begin VB.TextBox wood_ec_text 
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
               Left            =   2400
               TabIndex        =   22
               Text            =   "1"
               Top             =   330
               Width           =   1545
            End
            Begin VB.Label Label30 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   83
               Top             =   1800
               Width           =   705
            End
            Begin VB.Label Label29 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   82
               Top             =   1440
               Width           =   915
            End
            Begin VB.Label Label28 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   81
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label27 
               Caption         =   "m3/ton"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   4050
               TabIndex        =   80
               Top             =   720
               Width           =   885
            End
            Begin VB.Label Label26 
               Caption         =   "GJ/ton"
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
               Left            =   4050
               TabIndex        =   79
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               Caption         =   "NOx"
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
               Left            =   1170
               TabIndex        =   78
               Top             =   1770
               Width           =   1155
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "SO2"
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
               Left            =   1380
               TabIndex        =   77
               Top             =   1440
               Width           =   945
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "CO2"
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
               Left            =   1350
               TabIndex        =   76
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "Water"
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
               Left            =   1560
               TabIndex        =   75
               Top             =   690
               Width           =   765
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy Consuption"
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
               Left            =   450
               TabIndex        =   74
               Top             =   360
               Width           =   1875
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Agregates Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   1140
            TabIndex        =   62
            Top             =   7920
            Width           =   5865
            Begin VB.TextBox agregates_ec_text 
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
               Left            =   2400
               TabIndex        =   17
               Text            =   "1"
               Top             =   330
               Width           =   1545
            End
            Begin VB.TextBox agregates_water_text 
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
               Left            =   2400
               TabIndex        =   18
               Text            =   "1"
               Top             =   690
               Width           =   1545
            End
            Begin VB.TextBox agregates_co2_text 
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
               Left            =   2400
               TabIndex        =   19
               Text            =   "1"
               Top             =   1050
               Width           =   1545
            End
            Begin VB.TextBox agregates_so2_text 
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
               Left            =   2400
               TabIndex        =   20
               Text            =   "1"
               Top             =   1410
               Width           =   1545
            End
            Begin VB.TextBox agregates_nox_text 
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
               Left            =   2400
               TabIndex        =   21
               Text            =   "1"
               Top             =   1770
               Width           =   1545
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy Consuption"
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
               Left            =   450
               TabIndex        =   72
               Top             =   360
               Width           =   1875
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Water"
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
               Left            =   1560
               TabIndex        =   71
               Top             =   690
               Width           =   765
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               Caption         =   "CO2"
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
               Left            =   1350
               TabIndex        =   70
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "SO2"
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
               Left            =   1380
               TabIndex        =   69
               Top             =   1440
               Width           =   945
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "NOx"
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
               Left            =   1200
               TabIndex        =   68
               Top             =   1800
               Width           =   1155
            End
            Begin VB.Label Label15 
               Caption         =   "GJ/ton"
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
               Left            =   4050
               TabIndex        =   67
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label14 
               Caption         =   "m3/ton"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   4050
               TabIndex        =   66
               Top             =   720
               Width           =   885
            End
            Begin VB.Label Label13 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   65
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label12 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   64
               Top             =   1440
               Width           =   915
            End
            Begin VB.Label Label11 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   63
               Top             =   1800
               Width           =   705
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Cement Properties"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2235
            Left            =   1140
            TabIndex        =   51
            Top             =   5640
            Width           =   5865
            Begin VB.TextBox cement_nox_text 
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
               Left            =   2400
               TabIndex        =   16
               Text            =   "1"
               Top             =   1770
               Width           =   1545
            End
            Begin VB.TextBox cement_so2_text 
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
               Left            =   2400
               TabIndex        =   15
               Text            =   "1"
               Top             =   1410
               Width           =   1545
            End
            Begin VB.TextBox cement_co2_text 
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
               Left            =   2400
               TabIndex        =   14
               Text            =   "1"
               Top             =   1050
               Width           =   1545
            End
            Begin VB.TextBox cement_water_text 
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
               Left            =   2400
               TabIndex        =   13
               Text            =   "1"
               Top             =   690
               Width           =   1545
            End
            Begin VB.TextBox cement_ec_text 
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
               Left            =   2400
               TabIndex        =   12
               Text            =   "1"
               Top             =   330
               Width           =   1545
            End
            Begin VB.Label Label10 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   61
               Top             =   1800
               Width           =   705
            End
            Begin VB.Label Label9 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   60
               Top             =   1440
               Width           =   915
            End
            Begin VB.Label Label8 
               Caption         =   "Kg/ton"
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
               Left            =   4050
               TabIndex        =   59
               Top             =   1080
               Width           =   1035
            End
            Begin VB.Label Label7 
               Caption         =   "m3/ton"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   4050
               TabIndex        =   58
               Top             =   720
               Width           =   885
            End
            Begin VB.Label Label6 
               Caption         =   "GJ/ton"
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
               Left            =   4050
               TabIndex        =   57
               Top             =   360
               Width           =   885
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "NOx"
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
               Left            =   1200
               TabIndex        =   56
               Top             =   1800
               Width           =   1155
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "SO2"
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
               Left            =   1380
               TabIndex        =   55
               Top             =   1440
               Width           =   945
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "CO2"
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
               Left            =   1350
               TabIndex        =   54
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Water"
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
               Left            =   1560
               TabIndex        =   53
               Top             =   690
               Width           =   765
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Energy Consuption"
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
               Left            =   450
               TabIndex        =   52
               Top             =   360
               Width           =   1875
            End
         End
      End
   End
End
Attribute VB_Name = "Frm_database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim num_c_entries As Integer
Dim num_m_entries As Integer

Dim HeightDiff As Integer
Dim WidthDiff As Integer

Private Const FORMHEIGHT = 9000
Private Const FORMWIDTH = 10410

Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    check_db = True
End Sub


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

    SSTab.Tab = 0

    Call load_databases
    
    metalic_date_text = date
    concrete_date_text = date
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    On Error Resume Next
    If Me.Width < FORMWIDTH Then Me.Width = FORMWIDTH
    If Me.Height < FORMHEIGHT Then Me.Height = FORMHEIGHT
    If Err <> 0 Then
        Exit Sub
    End If
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

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'                   END OF GENERAL FORM FUNCTIONS
'                   start button click subs
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Private Sub concrete_copy_button_Click()
Dim last_index As Integer

last_index = concrete_entries.ListIndex
concrete_entries.ListIndex = num_c_entries
With concrete(last_index + 1)
    concrete_name_text.Text = .name
    concrete_date_text.Text = .date
    concrete_description_text.Text = .description
    With .wood
        wood_ec_text.Text = .energy
        wood_water_text.Text = .water
        wood_co2_text.Text = .co2
        wood_so2_text.Text = .so2
        wood_nox_text.Text = .nox
    End With
    With .cement
        cement_ec_text.Text = .energy
        cement_water_text.Text = .water
        cement_co2_text.Text = .co2
        cement_so2_text.Text = .so2
        cement_nox_text.Text = .nox
    End With
    With .steel
        steel_ec_text.Text = .energy
        steel_water_text.Text = .water
        steel_co2_text.Text = .co2
        steel_so2_text.Text = .so2
        steel_nox_text.Text = .nox
    End With
    With .agregates
        agregates_ec_text.Text = .energy
        agregates_water_text.Text = .water
        agregates_co2_text.Text = .co2
        agregates_so2_text.Text = .so2
        agregates_nox_text.Text = .nox
    End With
    With .water
        water_ec_text.Text = .energy
        water_water_text.Text = .water
        water_co2_text.Text = .co2
        water_so2_text.Text = .so2
        water_nox_text.Text = .nox
    End With
End With
    
End Sub
Private Sub metal_copy_button_Click()
Dim last_index As Integer

last_index = metalic_entries.ListIndex
metalic_entries.ListIndex = num_m_entries
With metalic(last_index + 1)
    metalic_name_text.Text = .name
    metalic_date_text.Text = .date
    metalic_description_text.Text = .description
    With .steel
        metalic_ec_text.Text = .energy
        metalic_water_text.Text = .water
        metalic_co2_text.Text = .co2
        metalic_so2_text.Text = .so2
        metalic_nox_text.Text = .nox
    End With
    With .transport
        metalic_distance_text.Text = .distance
        trans_co2_text.Text = .co2
        trans_so2_text.Text = .so2
        trans_nox_text.Text = .nox
    End With
End With
End Sub

Private Sub concrete_delete_button_Click()
Dim i As Integer
Dim filename As String
Dim chain As String
    
Err.Clear
filename = App.path & "\database\concrete.dbs"
On Error Resume Next
Kill filename
Open filename For Output As #1
If Err.Number <> 0 Then ' file not found?!
    MsgBox "Error Opening database file!", vbCritical + vbOKOnly, App.Title
    Exit Sub
End If
If UBound(concrete) > 2 Then
    Print #1, UBound(concrete) - 2
    For i = 1 To UBound(concrete) - 1
        If i <> concrete_entries.ListIndex + 1 Then
            chain = ""
            With concrete(i)
                chain = .name & "#" & .date & "#" & .description & "@"
                With .wood
                    chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
                End With
                With .cement
                    chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
                End With
                With .steel
                    chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
                End With
                With .agregates
                    chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
                End With
                With .water
                    chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
                End With
            End With
            Print #1, chain
        End If
    Next i
End If
Close #1
load_databases
End Sub


Private Sub metalic_delete_button_Click()
Dim i As Integer
Dim filename As String
Dim chain As String
    
Err.Clear
filename = App.path & "\database\steel.dbs"
On Error Resume Next
Kill filename
Open filename For Output As #1
If Err.Number <> 0 Then ' file not found?!
    MsgBox "Error Opening database file!", vbCritical + vbOKOnly, App.Title
    Exit Sub
End If
If UBound(metalic) > 2 Then
    Print #1, UBound(metalic) - 2
    For i = 1 To UBound(metalic) - 1
        If i <> metalic_entries.ListIndex + 1 Then
            chain = ""
            With metalic(i)
                chain = .name & "#" & .date & "#" & .description & "@"
                With .steel
                    chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
                End With
                With .transport
                    chain = chain & .co2 & "#" & .distance & "#" & .nox & "#" & .so2 & "@"
                End With
            End With
            Print #1, chain
        End If
    Next i
End If
Close #1
load_databases

End Sub

Private Sub concrete_entries_Click()
If concrete_entries.ListIndex = num_c_entries Then
        concrete_name_text.Text = " "
        concrete_date_text.Text = date
        concrete_description_text.Text = " "
        wood_ec_text.Text = " "
        wood_water_text.Text = " "
        wood_co2_text.Text = " "
        wood_so2_text.Text = " "
        wood_nox_text.Text = " "
        cement_ec_text.Text = " "
        cement_water_text.Text = " "
        cement_co2_text.Text = " "
        cement_so2_text.Text = " "
        cement_nox_text.Text = " "
        steel_ec_text.Text = " "
        steel_water_text.Text = " "
        steel_co2_text.Text = " "
        steel_so2_text.Text = " "
        steel_nox_text.Text = " "
        agregates_ec_text.Text = " "
        agregates_water_text.Text = " "
        agregates_co2_text.Text = " "
        agregates_so2_text.Text = " "
        agregates_nox_text.Text = " "
        water_ec_text.Text = " "
        water_water_text.Text = " "
        water_co2_text.Text = " "
        water_so2_text.Text = " "
        water_nox_text.Text = " "
Else
    Call load_entries("concrete", concrete_entries.ListIndex + 1)
End If
End Sub

Private Sub metalic_entries_Click()
If metalic_entries.ListIndex = num_m_entries Then
        metalic_name_text.Text = " "
        metalic_date_text.Text = date
        metalic_description_text.Text = " "
        metalic_ec_text.Text = " "
        metalic_water_text.Text = " "
        metalic_co2_text.Text = " "
        metalic_so2_text.Text = " "
        metalic_nox_text.Text = " "
        metalic_distance_text.Text = " "
        trans_co2_text.Text = " "
        trans_so2_text.Text = " "
        trans_nox_text.Text = " "
Else
    Call load_entries("metalic", metalic_entries.ListIndex + 1)
End If

End Sub

Private Sub metalic_save_button_Click()
If Not verify_data("metalic", metalic_entries.ListIndex + 1) Then
    Exit Sub
End If
Call save_data("metalic", metalic_entries.ListIndex + 1)
Call load_databases
End Sub

Private Sub concrete_save_button_Click()
If Not verify_data("concrete", concrete_entries.ListIndex + 1) Then
    Exit Sub
End If
Call save_data("concrete", concrete_entries.ListIndex + 1)
Call load_databases
End Sub

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'                  END OF BUTTON CLICK SUBS
'                  start of code support subs
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Private Sub load_entries(typ As String, s As Integer)
If typ = "concrete" Then
    With concrete(s)
        concrete_name_text.Text = .name
        concrete_date_text.Text = .date
        concrete_description_text.Text = .description
        With .wood
            wood_ec_text.Text = .energy
            wood_water_text.Text = .water
            wood_co2_text.Text = .co2
            wood_so2_text.Text = .so2
            wood_nox_text.Text = .nox
        End With
        With .cement
            cement_ec_text.Text = .energy
            cement_water_text.Text = .water
            cement_co2_text.Text = .co2
            cement_so2_text.Text = .so2
            cement_nox_text.Text = .nox
        End With
        With .steel
            steel_ec_text.Text = .energy
            steel_water_text.Text = .water
            steel_co2_text.Text = .co2
            steel_so2_text.Text = .so2
            steel_nox_text.Text = .nox
        End With
        With .agregates
            agregates_ec_text.Text = .energy
            agregates_water_text.Text = .water
            agregates_co2_text.Text = .co2
            agregates_so2_text.Text = .so2
            agregates_nox_text.Text = .nox
        End With
        With .water
            water_ec_text.Text = .energy
            water_water_text.Text = .water
            water_co2_text.Text = .co2
            water_so2_text.Text = .so2
            water_nox_text.Text = .nox
        End With
    End With
Else
    With metalic(s)
        metalic_name_text.Text = .name
        metalic_date_text.Text = .date
        metalic_description_text.Text = .description
        With .steel
            metalic_ec_text.Text = .energy
            metalic_water_text.Text = .water
            metalic_co2_text.Text = .co2
            metalic_so2_text.Text = .so2
            metalic_nox_text.Text = .nox
        End With
        With .transport
            metalic_distance_text.Text = .distance
            trans_co2_text.Text = .co2
            trans_so2_text.Text = .so2
            trans_nox_text.Text = .nox
        End With
    End With

End If
End Sub

Private Sub load_databases()

Dim i As Integer
Dim filename As String
Dim chain As String
Dim r() As String
Dim s() As String
Dim num As Integer

' loading concrete database
ReDim concrete(1)
Err.Clear
On Error Resume Next
filename = App.path & "\database\concrete.dbs"
Open filename For Input As #1
concrete_entries.Clear
If Err.Number = 0 Then ' file not found!?
    Input #1, num
    num_c_entries = num
    ReDim concrete(num + 1)
    i = 0
    While Not EOF(1)
        Input #1, chain
        i = i + 1
        r() = Split(chain, "@")
        s() = Split(r(0), "#")
        With concrete(i)
            .name = s(0)
            concrete_entries.AddItem .name
            .date = s(1)
            .description = s(2)
            s() = Split(r(1), "#")
            With .wood
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(2), "#")
            With .cement
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(3), "#")
            With .steel
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(4), "#")
            With .agregates
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(5), "#")
            With .water
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
        End With
    Wend
    
    If num = 1 Then
        num_concrete_entries.Caption = "There's " & CStr(num) & " entry in the concrete database"
    Else
        num_concrete_entries.Caption = "There are " & CStr(num) & " entries in the concrete database"
    End If
    concrete_entries.AddItem "New Entry"
    concrete_entries.ListIndex = num
Else
    num_concrete_entries.Caption = "There aren't any entries in the concrete database"
    concrete_entries.AddItem "New Entry"
    concrete_entries.ListIndex = 0
End If
Close #1
' loading metallic struct database
ReDim metalic(1)
Err.Clear
On Error Resume Next
filename = App.path & "\database\steel.dbs"
Open filename For Input As #1
metalic_entries.Clear
If Err.Number = 0 Then ' file not found!?
    Input #1, num
    num_m_entries = num
    ReDim metalic(num + 1)
    i = 0
    While Not EOF(1)
        Input #1, chain
        i = i + 1
        r() = Split(chain, "@")
        s() = Split(r(0), "#")
        With metalic(i)
            .name = s(0)
            metalic_entries.AddItem .name
            .date = s(1)
            .description = s(2)
            s() = Split(r(1), "#")
            With .steel
                .co2 = str2str(s(0))
                .energy = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
                .water = str2str(s(4))
            End With
            s() = Split(r(2), "#")
            With .transport
                .co2 = str2str(s(0))
                .distance = str2str(s(1))
                .nox = str2str(s(2))
                .so2 = str2str(s(3))
            End With
        End With
    Wend
    
    If num = 1 Then
        num_metalic_entries.Caption = "There's " & CStr(num) & " entry in the metalic database"
    Else
        num_metalic_entries.Caption = "There are " & CStr(num) & " entries in the metalic database"
    End If
    metalic_entries.AddItem "New Entry"
    metalic_entries.ListIndex = num
Else
    metalic_entries.AddItem "New Entry"
    metalic_entries.ListIndex = 0
    num_metalic_entries.Caption = "There aren't any entries in the metalic database"
End If
Close #1

End Sub

Private Function ifdata(ByRef s) As Boolean
If s = "" Or Not IsNumeric(s) Then
    ifdata = True
Else
    ifdata = False
End If
End Function

Function verify_data(where As String, Pos As Integer) As Boolean

Dim tmp As Boolean
verify_data = True
If where = "concrete" Then
    If concrete_name_text.Text = "" Then
        verify_data = False
        concrete_name_text.SetFocus
        Exit Function
    End If
    If concrete_date_text.Text = "" Then
        verify_data = False
        concrete_date_text.SetFocus
        Exit Function
    End If
    If concrete_description_text.Text = "" Then
        verify_data = False
        concrete_description_text.SetFocus
        Exit Function
    End If
    If ifdata(cement_ec_text.Text) Then
        verify_data = False
        cement_ec_text.SetFocus
        Exit Function
    End If
    If ifdata(cement_water_text.Text) Then
        verify_data = False
        cement_water_text.SetFocus
        Exit Function
    End If
    If ifdata(cement_co2_text.Text) Then
        verify_data = False
        cement_co2_text.SetFocus
        Exit Function
    End If
    If ifdata(cement_so2_text.Text) Then
        verify_data = False
        cement_so2_text.SetFocus
        Exit Function
    End If
    If ifdata(cement_nox_text.Text) Then
        verify_data = False
        cement_nox_text.SetFocus
        Exit Function
    End If
    
    If ifdata(wood_ec_text.Text) Then
        verify_data = False
        wood_ec_text.SetFocus
        Exit Function
    End If
    If ifdata(wood_water_text.Text) Then
        verify_data = False
        wood_water_text.SetFocus
        Exit Function
    End If
    If ifdata(wood_co2_text.Text) Then
        verify_data = False
        wood_co2_text.SetFocus
        Exit Function
    End If
    If ifdata(wood_so2_text.Text) Then
        verify_data = False
        wood_so2_text.SetFocus
        Exit Function
    End If
    If ifdata(wood_nox_text.Text) Then
        verify_data = False
        wood_nox_text.SetFocus
        Exit Function
    End If
    
    If ifdata(steel_ec_text.Text) Then
        verify_data = False
        steel_ec_text.SetFocus
        Exit Function
    End If
    If ifdata(steel_water_text.Text) Then
        verify_data = False
        steel_water_text.SetFocus
        Exit Function
    End If
    If ifdata(steel_co2_text.Text) Then
        verify_data = False
        steel_co2_text.SetFocus
        Exit Function
    End If
    If ifdata(steel_so2_text.Text) Then
        verify_data = False
        steel_so2_text.SetFocus
        Exit Function
    End If
    If ifdata(steel_nox_text.Text) Then
        verify_data = False
        steel_nox_text.SetFocus
        Exit Function
    End If
    
    If ifdata(agregates_ec_text.Text) Then
        verify_data = False
        agregates_ec_text.SetFocus
        Exit Function
    End If
    If ifdata(agregates_water_text.Text) Then
        verify_data = False
        agregates_water_text.SetFocus
        Exit Function
    End If
    If ifdata(agregates_co2_text.Text) Then
        verify_data = False
        agregates_co2_text.SetFocus
        Exit Function
    End If
    If ifdata(agregates_so2_text.Text) Then
        verify_data = False
        agregates_so2_text.SetFocus
        Exit Function
    End If
    If ifdata(agregates_nox_text.Text) Then
        verify_data = False
        agregates_nox_text.SetFocus
        Exit Function
    End If
    
    If ifdata(water_ec_text.Text) Then
        verify_data = False
        water_ec_text.SetFocus
        Exit Function
    End If
    If ifdata(water_water_text.Text) Then
        verify_data = False
        water_water_text.SetFocus
        Exit Function
    End If
    If ifdata(water_co2_text.Text) Then
        verify_data = False
        water_co2_text.SetFocus
        Exit Function
    End If
    If ifdata(water_so2_text.Text) Then
        verify_data = False
        water_so2_text.SetFocus
        Exit Function
    End If
    If ifdata(water_nox_text.Text) Then
        verify_data = False
        water_nox_text.SetFocus
        Exit Function
    End If
    
    With concrete(Pos)
        .name = concrete_name_text.Text
        .date = concrete_date_text.Text
        .description = concrete_description_text.Text
        With .agregates
            .co2 = agregates_co2_text.Text
            .energy = agregates_ec_text.Text
            .nox = agregates_nox_text.Text
            .so2 = agregates_so2_text.Text
            .water = agregates_water_text.Text
        End With
        With .wood
            .co2 = wood_co2_text.Text
            .energy = wood_ec_text.Text
            .nox = wood_nox_text.Text
            .so2 = wood_so2_text.Text
            .water = wood_water_text.Text
        End With
        With .steel
            .co2 = steel_co2_text.Text
            .energy = steel_ec_text.Text
            .nox = steel_nox_text.Text
            .so2 = steel_so2_text.Text
            .water = steel_water_text.Text
        End With
        With .cement
            .co2 = cement_co2_text.Text
            .energy = cement_ec_text.Text
            .nox = cement_nox_text.Text
            .so2 = cement_so2_text.Text
            .water = cement_water_text.Text
        End With
        With .water
            .co2 = water_co2_text.Text
            .energy = water_ec_text.Text
            .nox = water_nox_text.Text
            .so2 = water_so2_text.Text
            .water = water_water_text.Text
        End With
    End With
    
Else ' metalic
    If metalic_name_text.Text = "" Then
        verify_data = False
        metalic_name_text.SetFocus
        Exit Function
    End If
    If metalic_date_text.Text = "" Then
        verify_data = False
        metalic_date_text.SetFocus
        Exit Function
    End If
    If metalic_description_text.Text = "" Then
        verify_data = False
        metalic_description_text.SetFocus
        Exit Function
    End If

    If ifdata(metalic_ec_text.Text) Then
        verify_data = False
        metalic_ec_text.SetFocus
        Exit Function
    End If
    If ifdata(metalic_water_text.Text) Then
        verify_data = False
        metalic_water_text.SetFocus
        Exit Function
    End If
    If ifdata(metalic_co2_text.Text) Then
        verify_data = False
        metalic_co2_text.SetFocus
        Exit Function
    End If
    If ifdata(metalic_so2_text.Text) Then
        verify_data = False
        metalic_so2_text.SetFocus
        Exit Function
    End If
    If ifdata(metalic_nox_text.Text) Then
        verify_data = False
        metalic_nox_text.SetFocus
        Exit Function
    End If

    
    If ifdata(metalic_distance_text.Text) Then
        verify_data = False
        metalic_distance_text.SetFocus
        Exit Function
    End If
    If ifdata(trans_co2_text.Text) Then
        verify_data = False
        trans_co2_text.SetFocus
        Exit Function
    End If
    If ifdata(trans_so2_text.Text) Then
        verify_data = False
        trans_so2_text.SetFocus
        Exit Function
    End If
    If ifdata(trans_nox_text.Text) Then
        verify_data = False
        trans_nox_text.SetFocus
        Exit Function
    End If
    With metalic(Pos)
        .name = metalic_name_text.Text
        .date = metalic_date_text.Text
        .description = metalic_description_text.Text
        With .steel
            .co2 = convert_type(metalic_co2_text.Text)
            .energy = convert_type(metalic_ec_text.Text)
            .nox = convert_type(metalic_nox_text.Text)
            .so2 = convert_type(metalic_so2_text.Text)
            .water = convert_type(metalic_water_text.Text)
        End With
        With .transport
            .co2 = convert_type(trans_co2_text.Text)
            .distance = convert_type(metalic_distance_text.Text)
            .nox = convert_type(trans_nox_text.Text)
            .so2 = convert_type(trans_so2_text.Text)
        End With
    End With
End If
End Function

Sub save_data(s As String, Pos As Integer)

Dim i As Integer
Dim filename As String
Dim chain As String
Dim top As Integer

If s = "concrete" Then
    Err.Clear
    filename = App.path & "\database\concrete.dbs"
    On Error Resume Next
    Kill filename
    Open filename For Output As #1
    If Err.Number <> 0 Then ' file not found?!
        MsgBox "Error Opening database file!", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    If Pos = num_c_entries + 1 Then
        top = num_c_entries + 1
    Else
        top = num_c_entries
    End If
    Print #1, top
    For i = 1 To top
        chain = ""
        With concrete(i)
            chain = .name & "#" & .date & "#" & .description & "@"
            With .wood
                chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
            End With
            With .cement
                chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
            End With
            With .steel
                chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
            End With
            With .agregates
                chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
            End With
            With .water
                chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
            End With
        End With
        chain = Replace(chain, ",", ".")
        Print #1, chain
    Next i
    Close #1
Else
    Err.Clear
    filename = App.path & "\database\steel.dbs"
    On Error Resume Next
    Kill filename
    Open filename For Output As #1
    If Err.Number <> 0 Then ' file not found?!
        MsgBox "Error Opening database file!", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    If Pos = num_m_entries + 1 Then
        top = num_m_entries + 1
    Else
        top = num_m_entries
    End If
    Print #1, top
    For i = 1 To top
        chain = ""
        With metalic(i)
            chain = .name & "#" & .date & "#" & .description & "@"
            With .steel
                chain = chain & .co2 & "#" & .energy & "#" & .nox & "#" & .so2 & "#" & .water & "@"
            End With
            With .transport
                chain = chain & .co2 & "#" & .distance & "#" & .nox & "#" & .so2 & "@"
            End With
        End With
        chain = Replace(chain, ",", ".")
        Print #1, chain
    Next i
    Close #1
End If
End Sub

Private Sub Timer1_Timer()

If concrete_entries.Text = "New Entry" Then
    If concrete_copy_button.Enabled = True Then
        concrete_copy_button.Enabled = False
        concrete_delete_button.Enabled = False
    End If
Else
    If concrete_copy_button.Enabled = False Then
        concrete_copy_button.Enabled = True
        concrete_delete_button.Enabled = True
    End If
End If
If metalic_entries.Text = "New Entry" Then
    If metal_copy_button.Enabled = True Then
        metal_copy_button.Enabled = False
        metalic_delete_button.Enabled = False
    End If
Else
    If metal_copy_button.Enabled = False Then
        metal_copy_button.Enabled = True
        metalic_delete_button.Enabled = True
    End If
End If


End Sub
