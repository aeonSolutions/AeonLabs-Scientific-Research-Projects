VERSION 5.00
Begin VB.Form frm_choice 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EcoBuild - General Information"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   Icon            =   "frm_choice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
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
      Left            =   5430
      TabIndex        =   4
      Top             =   1830
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
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
      Left            =   7050
      TabIndex        =   1
      Top             =   1830
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Caption         =   "Structure"
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
      Left            =   1260
      TabIndex        =   5
      Top             =   750
      Width           =   6195
      Begin VB.OptionButton metallic_button 
         Caption         =   "Metallic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton concrete_button 
         Caption         =   "Concrete"
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
         Left            =   3420
         TabIndex        =   3
         Top             =   300
         Width           =   2505
      End
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   60
      Picture         =   "frm_choice.frx":324A
      Top             =   60
      Width           =   960
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   7800
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the structure type:"
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
      Left            =   1110
      TabIndex        =   0
      Top             =   210
      Width           =   4035
   End
End
Attribute VB_Name = "frm_choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim doc As Integer

doc = current_form()

frm_choice.Hide
If concrete_button.Value Then
    frm_concrete_board.Show 1
ElseIf metallic_button.Value Then
    frm_metallic_board.Show 1
End If
Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call DisableX(frm_choice)
End Sub

