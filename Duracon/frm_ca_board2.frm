VERSION 5.00
Begin VB.Form frm_ca_board2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DURACON - Profile Parameters"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7935
   Icon            =   "frm_ca_board2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton back_button 
      Caption         =   "лл Back"
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
      Left            =   3360
      TabIndex        =   61
      Top             =   5520
      Width           =   1425
   End
   Begin VB.CommandButton close_button 
      Caption         =   "Close"
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
      Left            =   4890
      TabIndex        =   60
      Top             =   5520
      Width           =   1425
   End
   Begin VB.CommandButton next_button 
      Caption         =   "Next "
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
      Left            =   6420
      TabIndex        =   29
      Top             =   5520
      Width           =   1425
   End
   Begin VB.Frame Title_4 
      Caption         =   "Profile Values"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   150
      TabIndex        =   33
      Top             =   2970
      Width           =   7665
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   4
         Left            =   2520
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   5
         Left            =   3120
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   6
         Left            =   3720
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   7
         Left            =   4320
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   8
         Left            =   4920
         TabIndex        =   25
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   9
         Left            =   5520
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   10
         Left            =   6120
         TabIndex        =   27
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   11
         Left            =   6720
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_c 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   1680
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   11
         Left            =   6750
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   8
         Left            =   4920
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   9
         Left            =   5520
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   10
         Left            =   6120
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   4
         Left            =   2520
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   5
         Left            =   3120
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   6
         Left            =   3720
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   7
         Left            =   4320
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.TextBox PrfVal_d 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6720
         TabIndex        =   59
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   58
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   57
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   56
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   55
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   54
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   53
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   51
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   50
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   49
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   " Chloride concentration (% weight of cement or concrete)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label4 
         Caption         =   " Depth from concrete surface (cm)"
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
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6720
         TabIndex        =   45
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6120
         TabIndex        =   44
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5520
         TabIndex        =   43
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4920
         TabIndex        =   42
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3720
         TabIndex        =   40
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3120
         TabIndex        =   39
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   38
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   37
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   720
         TabIndex        =   35
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.Frame Title_3 
      Caption         =   "Profile Data "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   31
      Top             =   1860
      Width           =   7695
      Begin VB.ComboBox Prfsize 
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
         ItemData        =   "frm_ca_board2.frx":324A
         Left            =   4470
         List            =   "frm_ca_board2.frx":3266
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   " Number of measurements for the profile [5-12]"
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
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   4365
      End
   End
   Begin VB.Frame Title_2 
      Caption         =   "Profile Prediction "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7695
      Begin VB.TextBox Age 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2070
            SubFormatType   =   0
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
         Left            =   2160
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Prfpred_2 
         Caption         =   "Yes      Time (days)"
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
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Prfpred_1 
         Caption         =   "No"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Label Title_1 
      Caption         =   "   PROFILE INFORMATION"
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
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frm_ca_board2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim doc As Integer

Private Sub Age_Change()
Dim tmp() As String
Dim tmp2 As String

With doc_props(doc)
    If Age.Text <> "" Then
        tmp2 = Age.Text
        tmp() = Split(Age.Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            tmp2 = "0" & Age.Text
        End If
        tmp() = Split(Age.Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            tmp2 = "0" & Age.Text
        End If
        If IsNumeric(tmp2) Then
            On Error Resume Next
            .predage = tmp2
            Prfpred_2.Value = True
            Prfpred_1.Value = False
            If Err.Number <> 0 Then
                If Err.Number = 6 Then
                    'Numeric value too large
                    Age.SetFocus
                    Exit Sub
                Else
                    'Please report this bug: " & "Error " & Err.Number
                    Age.SetFocus
                    Exit Sub
                End If
            End If
        End If
    Else
        'Only Positive numeric Values allowed!
        Age.SetFocus
        Exit Sub
    End If
End With
End Sub

Private Sub back_button_Click()

If store_data(False) Then doc_props(doc).frm_ca_board2_values.ready = False
If last_window = "frm_ca_board1" Then
    frm_ca_board2.Hide
    frm_ca_board1.Show 1
    Unload Me
ElseIf last_window = "frm_ca_board2" Then
    frm_ca_board2.Hide
    frm_ca_board2.Show 1
    Unload Me
End If

End Sub
Private Function store_data(msg As Boolean) As Boolean
Dim j As Integer
Dim i As Integer
Dim Index As Integer
Dim var1 As Single
Dim var2 As Single
Dim tmp() As String

store_data = True
With doc_props(doc)
    If Age.Text <> "" Then
        tmp() = Split(Age.Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            Age.Text = "0" & Age.Text
        End If
        tmp() = Split(Age.Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            Age.Text = "0" & Age.Text
        End If
        If IsNumeric(Age.Text) Then
            On Error Resume Next
            .predage = Age.Text
            Prfpred_2.Value = True
            Prfpred_1.Value = False
            If Err.Number <> 0 Then
                If Err.Number = 6 Then
                    MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                    Age.SetFocus
                    store_data = False
                    Exit Function
                Else
                    MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOK + vbCritical, "Duracon"
                    Age.SetFocus
                    store_data = False
                    Exit Function
                End If
            End If
        End If
    Else
        If Prfpred_2.Value = True Then
            MsgBox "Only Positive numeric Values allowed!", vbOK + vbCritical, "Duracon"
            Age.SetFocus
            store_data = False
            Exit Function
        End If
    End If

For Index = 0 To 11
        If PrfVal_d(Index).Text <> "" Then
            tmp() = Split(PrfVal_d(Index).Text, ".")
            If tmp(0) = "." Or tmp(0) = "" Then
                PrfVal_d(Index).Text = "0" & PrfVal_d(Index).Text
            End If
            tmp() = Split(PrfVal_d(Index).Text, ",")
            If tmp(0) = "," Or tmp(0) = "" Then
                PrfVal_d(Index).Text = "0" & PrfVal_d(Index).Text
            End If
            If IsNumeric(PrfVal_d(Index).Text) Then
                On Error Resume Next
                .prfvdp(Index) = PrfVal_d(Index).Text
                If Err.Number <> 0 Then
                    If Err.Number = 6 Then
                        MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                        PrfVal_d(Index).SetFocus
                        store_data = False
                        Exit Function
                    Else
                        MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOK + vbCritical, "Duracon"
                        PrfVal_d(Index).SetFocus
                        store_data = False
                        Exit Function
                    End If
                End If
            End If
        Else
            MsgBox "Only Positive numeric Values allowed!", vbOK + vbCritical, "Duracon"
            PrfVal_d(Index).SetFocus
            store_data = False
            Exit Function
        End If
        
        If PrfVal_c(Index).Text <> "" Then
            tmp() = Split(PrfVal_c(Index).Text, ".")
            If tmp(0) = "." Or tmp(0) = "" Then
                PrfVal_c(Index).Text = "0" & PrfVal_c(Index).Text
            End If
            tmp() = Split(PrfVal_c(Index).Text, ",")
            If tmp(0) = "," Or tmp(0) = "" Then
                PrfVal_c(Index).Text = "0" & PrfVal_c(Index).Text
            End If
            If IsNumeric(PrfVal_c(Index).Text) Then
                On Error Resume Next
                .prfvcs(Index) = PrfVal_c(Index).Text
                If Err.Number <> 0 Then
                    If Err.Number = 6 Then
                        MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                        PrfVal_c(Index).SetFocus
                        store_data = False
                        Exit Function
                    Else
                        MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOK + vbCritical, "Duracon"
                        PrfVal_c(Index).SetFocus
                        store_data = False
                        Exit Function
                    End If
                End If
            End If
        Else
            MsgBox "Only Positive numeric Values allowed!", vbOK + vbCritical, "Duracon"
            PrfVal_c(Index).SetFocus
            Exit Function
        End If
Next Index
End With
store_data = True
With doc_props(doc)
    With .frm_ca_board2_values
        If Age.Text <> "" Then
            .Age = Age.Text
            doc_props(doc).predage = .Age
        Else
            .Age = "N/A"
            doc_props(doc).predage = 0
        End If
        .Prfsize = Prfsize.ListIndex
    End With
    
    For j = 0 To 11
        .prfvdp(j) = Val(PrfVal_d(j).Text)
        .prfvcs(j) = Val(PrfVal_c(j).Text)
    Next j
    If .prfsz < 5 Then
        If msg Then MsgBox "Please Fill in all cells", vbOK + vbCritical, "Duracon"

        store_data = False
    Else
        For i = 0 To (.prfsz - 2)
            If (.prfvdp(i) >= .prfvdp(i + 1)) Then
                If msg Then MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                store_data = False
                Exit Function
            End If
        Next i
    End If
End With

End Function

Private Sub next_button_Click()
Dim i As Integer
Dim j As Integer

doc_props(doc).frm_ca_board2_values.values = True
last_window = "frm_ca_board2"
If store_data(True) Then
    frm_ca_board2.Hide
    frm_ca_board3.Show 1
    Unload Me
End If
doc_props(doc).frm_ca_board2_values.ready = False
End Sub

Private Sub close_button_Click()
    FState(doc).values = True
    doc_props(doc).frm_ca_board2_values.values = True
    If store_data(False) Then doc_props(doc).frm_ca_board2_values.ready = False
    Call refresh_lista(doc)
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer

Call DisableX(frm_ca_board2)
doc = current_form
With doc_props(doc)
    With .frm_ca_board2_values
        If .values Then
            If .Age = "N/A" Then
                Age.Text = ""
            Else
                Age.Text = .Age
            End If
            Prfsize.ListIndex = .Prfsize
            For i = 0 To 4 + .Prfsize
                PrfVal_d(i).Enabled = True
                PrfVal_c(i).Enabled = True
            Next i
            If .Prfsize < 8 Then
                For i = 5 + .Prfsize To 11
                    PrfVal_d(i).Enabled = False
                    PrfVal_c(i).Enabled = False
                Next i
            End If
        Else
            .Age = "N/A"
            .Prfsize = 1
            Prfsize.ListIndex = 0
        End If
    End With
        For i = 0 To 11
            PrfVal_d(i).Text = .prfvdp(i)
            PrfVal_c(i).Text = .prfvcs(i)
        Next i
End With

End Sub

Private Sub Prfpred_1_Click()

    doc_props(doc).prfpred = 0
    doc_props(doc).frm_ca_board2_values.Age = "N/A"
    Age.Text = ""
    Prfpred_2.Value = False
    Prfpred_1.Value = True
End Sub

Private Sub Prfpred_2_Click()
    doc_props(doc).prfpred = 1
End Sub


Private Sub Prfsize_Click()
Dim i As Integer

If Prfsize.Text <> "" Then
    doc_props(doc).prfsz = Val(Prfsize.Text)
End If
For i = 0 To doc_props(doc).prfsz - 1
    PrfVal_d(i).Enabled = True
    PrfVal_c(i).Enabled = True
Next i
If doc_props(doc).prfsz < 12 Then
    For i = doc_props(doc).prfsz To 11
        PrfVal_d(i).Enabled = False
        PrfVal_c(i).Enabled = False
    Next i
End If
End Sub

