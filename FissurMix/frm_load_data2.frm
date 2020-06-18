VERSION 5.00
Begin VB.Form frm_load_data2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[CODED]"
   ClientHeight    =   5340
   ClientLeft      =   3345
   ClientTop       =   2505
   ClientWidth     =   7485
   Icon            =   "frm_load_data2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Statistic Analysis Parameters"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   22
      Top             =   3420
      Width           =   7215
      Begin VB.TextBox sl_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   5400
         TabIndex        =   47
         Top             =   750
         Width           =   825
      End
      Begin VB.TextBox elements_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   2430
         TabIndex        =   29
         Top             =   300
         Width           =   825
      End
      Begin VB.TextBox m_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
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
         Left            =   4740
         TabIndex        =   8
         Text            =   "6"
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox s0_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   750
         Width           =   825
      End
      Begin VB.TextBox rs_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   1530
         TabIndex        =   6
         Top             =   780
         Width           =   825
      End
      Begin VB.Label Label30 
         Caption         =   "L"
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
         Left            =   5220
         TabIndex        =   49
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label29 
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
         Height          =   255
         Left            =   6270
         TabIndex        =   48
         Top             =   780
         Width           =   405
      End
      Begin VB.Label Label17 
         Caption         =   "Number of elements"
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
         Left            =   600
         TabIndex        =   30
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label16 
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
         Height          =   255
         Left            =   4230
         TabIndex        =   28
         Top             =   780
         Width           =   405
      End
      Begin VB.Label Label15 
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
         Height          =   255
         Left            =   2400
         TabIndex        =   27
         Top             =   810
         Width           =   405
      End
      Begin VB.Label Label14 
         Caption         =   "0"
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
         Left            =   3180
         TabIndex        =   26
         Top             =   870
         Width           =   225
      End
      Begin VB.Label Label13 
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
         Height          =   225
         Left            =   4500
         TabIndex        =   25
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   24
         Top             =   780
         Width           =   165
      End
      Begin VB.Label scale_info_2 
         Caption         =   "Residual Stress"
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
         Left            =   150
         TabIndex        =   23
         Top             =   810
         Width           =   1485
      End
      Begin VB.Label Label31 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5100
         TabIndex        =   50
         Top             =   780
         Width           =   165
      End
   End
   Begin VB.CommandButton finish_button 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5940
      TabIndex        =   11
      Top             =   4800
      Width           =   1365
   End
   Begin VB.CommandButton Cancel_button 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   4800
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample Phisical Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   90
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.TextBox esr_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
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
         Left            =   5370
         TabIndex        =   40
         Top             =   2340
         Width           =   825
      End
      Begin VB.TextBox efr_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
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
         Left            =   2190
         TabIndex        =   37
         Top             =   2040
         Width           =   825
      End
      Begin VB.TextBox es_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
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
         Left            =   5370
         TabIndex        =   34
         Top             =   1980
         Width           =   825
      End
      Begin VB.TextBox width_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   5370
         TabIndex        =   31
         Top             =   510
         Width           =   825
      End
      Begin VB.TextBox ef_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
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
         Left            =   2190
         TabIndex        =   3
         Top             =   1680
         Width           =   825
      End
      Begin VB.TextBox ts_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   5370
         TabIndex        =   5
         Top             =   1620
         Width           =   825
      End
      Begin VB.TextBox substrate_pc_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   5370
         TabIndex        =   4
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox tf_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   2190
         TabIndex        =   2
         Top             =   1320
         Width           =   825
      End
      Begin VB.TextBox lenght_txt 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
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
         Left            =   2190
         TabIndex        =   1
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label11 
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
         Height          =   225
         Left            =   6360
         TabIndex        =   51
         Top             =   1650
         Width           =   255
      End
      Begin VB.Label Label28 
         Caption         =   "r"
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
         Left            =   2040
         TabIndex        =   46
         Top             =   1980
         Width           =   135
      End
      Begin VB.Label Label27 
         Caption         =   "r"
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
         Left            =   5220
         TabIndex        =   45
         Top             =   2280
         Width           =   135
      End
      Begin VB.Label Label26 
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5100
         TabIndex        =   44
         Top             =   2310
         Width           =   165
      End
      Begin VB.Label Label25 
         Caption         =   "e"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   12.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   43
         Top             =   2010
         Width           =   165
      End
      Begin VB.Label Label24 
         Caption         =   "%"
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
         Left            =   6270
         TabIndex        =   42
         Top             =   2370
         Width           =   405
      End
      Begin VB.Label Label23 
         Caption         =   "s"
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
         Left            =   5190
         TabIndex        =   41
         Top             =   2520
         Width           =   135
      End
      Begin VB.Label Label22 
         Caption         =   "%"
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
         Left            =   3090
         TabIndex        =   39
         Top             =   2070
         Width           =   405
      End
      Begin VB.Label Label21 
         Caption         =   "f"
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
         Left            =   2040
         TabIndex        =   38
         Top             =   2250
         Width           =   135
      End
      Begin VB.Label Label20 
         Caption         =   "GPa"
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
         Left            =   6270
         TabIndex        =   36
         Top             =   2040
         Width           =   405
      End
      Begin VB.Label Label19 
         Caption         =   "Es"
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
         Left            =   5100
         TabIndex        =   35
         Top             =   2010
         Width           =   195
      End
      Begin VB.Label Label18 
         Caption         =   "Width"
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
         Left            =   4800
         TabIndex        =   33
         Top             =   540
         Width           =   705
      End
      Begin VB.Label Label8 
         Caption         =   "mm"
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
         Left            =   6240
         TabIndex        =   32
         Top             =   540
         Width           =   345
      End
      Begin VB.Label scale_info_1 
         Caption         =   "Poisson Coef."
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
         Left            =   4110
         TabIndex        =   21
         Top             =   1320
         Width           =   1305
      End
      Begin VB.Label Label12 
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
         Height          =   225
         Left            =   6240
         TabIndex        =   20
         Top             =   1650
         Width           =   165
      End
      Begin VB.Label Label10 
         Caption         =   "Ef"
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
         Left            =   1950
         TabIndex        =   19
         Top             =   1740
         Width           =   195
      End
      Begin VB.Label Label7 
         Caption         =   "Substrate:"
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
         Left            =   3840
         TabIndex        =   18
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "mm"
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
         Left            =   3060
         TabIndex        =   17
         Top             =   510
         Width           =   345
      End
      Begin VB.Label Label5 
         Caption         =   "nm"
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
         Left            =   3060
         TabIndex        =   16
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label4 
         Caption         =   "GPa"
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
         Left            =   3090
         TabIndex        =   15
         Top             =   1710
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "Ts"
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
         Left            =   5100
         TabIndex        =   14
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Tf"
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
         Left            =   1950
         TabIndex        =   13
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label num_concrete_entries 
         Caption         =   "Lenght"
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
         Left            =   1560
         TabIndex        =   12
         Top             =   510
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Film:"
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
         Left            =   1560
         TabIndex        =   9
         Top             =   1050
         Width           =   495
      End
   End
End
Attribute VB_Name = "frm_load_data2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private doc As Integer


Private Sub Cancel_button_Click()
    Call refresh_richtext
    Call reset_props
    Me.Hide
    Unload Me
End Sub

Private Sub scale_combo_Click()
    scale_info_1.Caption = scale_Combo.Text
    scale_info_2.Caption = scale_Combo.Text
    With doc_props(doc)
        With .exp_data
            .scales = scale_Combo.Text
        End With
    End With
End Sub

Private Sub Form_Load()
    doc = current_form
    Call DisableX(frm_load_data2)
    Me.Caption = App.Title
    If FState(doc).values Then
        With doc_props(doc).phisical
            width_txt.Text = .width_ * 1000 ' converted from m to mm
            lenght_txt.Text = .lenght * 1000 ' converted from m to mm
            ef_txt.Text = .ef / 1000000000# ' converted from Pa to GPa
            es_txt.Text = .es / 1000000000# ' converted from Pa to GPa
            efr_txt.Text = .efr * 100# ' converted to %
            esr_txt.Text = .esr * 100# ' converted to %
            substrate_pc_txt.Text = .substrate_pc
            tf_txt.Text = .tf * 1000000000 ' converted from m to nm
            ts_txt.Text = .ts * 1000000 ' converted from m to um
        End With
        With doc_props(doc).statistic
            m_txt.Text = .m
            rs_txt.Text = .rs / 1000000# ' converted from Pa to MPa
            s0_txt.Text = .s0 / 1000000# ' converted from Pa to MPa
            sl_txt.Text = .sl / 1000000# ' converted from Pa to MPa
            elements_txt.Text = .elements
        End With
    End If
End Sub
Private Sub finish_button_Click()
Dim area, multiply As Double
Dim c_term As Double

    If Not IsNumeric(lenght_txt.Text) Then
        lenght_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(substrate_pc_txt.Text) Then
        substrate_pc_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(tf_txt.Text) Then
        tf_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(ts_txt.Text) Then
        ts_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(ef_txt.Text) Then
        ef_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(es_txt.Text) Then
        es_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(efr_txt.Text) Then
        efr_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(esr_txt.Text) Then
        esr_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(rs_txt.Text) Then
        rs_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(s0_txt.Text) Then
        s0_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(sl_txt.Text) Then
        sl_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(m_txt.Text) Then
        m_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(elements_txt.Text) Then
        elements_txt.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(width_txt.Text) Then
        area_txt.SetFocus
        Exit Sub
    End If
    
    With doc_props(doc)
        With .phisical
            .width_ = convert_type(width_txt.Text) / 1000 ' converted from mm to m
            .lenght = convert_type(lenght_txt.Text) / 1000 ' converted from mm to m
            .substrate_pc = convert_type(substrate_pc_txt.Text)
            .tf = convert_type(tf_txt.Text) / 1000000000 ' converted from nm to m
            .ts = convert_type(ts_txt.Text) / 1000000 ' converted from um to m
            .ef = convert_type(ef_txt.Text) * 1000000000# ' converted from GPa to Pa
            .es = convert_type(es_txt.Text) * 1000000000# ' converted from GPa to Pa
            .efr = convert_type(efr_txt.Text) / 100 ' converted from % to adimensional
            .esr = convert_type(esr_txt.Text) / 100 ' converted from % to adimensional
            
            .area_s = .ts * .width_
            .area_f = .tf * .width_
            .area_total = .area_s + .area_f
        End With
        With .statistic
            .m = convert_type(m_txt.Text)
            .rs = convert_type(rs_txt.Text) * 1000000# ' converted from MPa to Pa
            .s0 = convert_type(s0_txt.Text) * 1000000# ' converted from MPa to Pa
            .sl = convert_type(sl_txt.Text) * 1000000# ' converted from MPa to Pa
            .elements = convert_type(elements_txt.Text)
        End With
        
        area = .phisical.area_total * 1000000 ' converted from m2 to mm2
        If .exp_data.scales = "MN" Then
            multiply = 1000000
        ElseIf .exp_data.scales = "KN" Then
            multiply = 1000
        ElseIf .exp_data.scales = "N" Then
            multiply = 1
        End If
        With .stress_c_curve
            .x5 = doc_props(doc).exp_data.x5 * multiply / area * 1000000# ' N/mm2 = MPa convertido de MPa para Pa
            .x4 = doc_props(doc).exp_data.x4 * multiply / area * 1000000#
            .x3 = doc_props(doc).exp_data.x3 * multiply / area * 1000000#
            .x2 = doc_props(doc).exp_data.x2 * multiply / area * 1000000#
            .x = doc_props(doc).exp_data.x * multiply / area * 1000000#
            .c = doc_props(doc).exp_data.c * multiply / area * 1000000#
        End With
                
        With .modulus_c_curve ' func already in Pa
            .x4 = doc_props(doc).stress_c_curve.x5 * 5
            .x3 = doc_props(doc).stress_c_curve.x4 * 4
            .x2 = doc_props(doc).stress_c_curve.x3 * 3
            .x = doc_props(doc).stress_c_curve.x2 * 2
            .c = doc_props(doc).stress_c_curve.x
        End With
        
        With .phisical
            c_term = (.ts + .tf) / (.ts * .ef)
        End With
    End With
    doc_props(doc).elements_generated = False
    FState(doc).saved = False
    FState(doc).values = True
    Call refresh_richtext

    Me.Hide
    Unload Me
End Sub

