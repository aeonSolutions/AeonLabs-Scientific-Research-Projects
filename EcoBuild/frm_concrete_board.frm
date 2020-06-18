VERSION 5.00
Begin VB.Form frm_concrete_board 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EcoBuild - Concrete data"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7575
   Icon            =   "frm_concrete_board.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Next_button 
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
      Left            =   5970
      TabIndex        =   8
      Top             =   3720
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
      Left            =   4470
      TabIndex        =   7
      Top             =   3720
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database selection"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   7215
      Begin VB.ComboBox concrete_combo 
         Height          =   315
         Left            =   3210
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   900
         Width           =   2535
      End
      Begin VB.Label num_concrete_entries 
         Caption         =   "[CODED]"
         Height          =   555
         Left            =   300
         TabIndex        =   9
         Top             =   360
         Width           =   6645
      End
      Begin VB.Label Label5 
         Caption         =   "Select a database:"
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
         Left            =   1470
         TabIndex        =   6
         Top             =   930
         Width           =   1695
      End
      Begin VB.Label name_txt 
         Caption         =   "[DB Name] - CODED"
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
         Left            =   900
         TabIndex        =   4
         Top             =   1620
         Width           =   5985
      End
      Begin VB.Label description_txt 
         Caption         =   "[DB Description] - CODED"
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
         Left            =   1350
         TabIndex        =   3
         Top             =   1980
         Width           =   5625
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
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
         TabIndex        =   2
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Description:"
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
         TabIndex        =   1
         Top             =   1980
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_concrete_board"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private concrete() As concrete_type

Private Sub Cancel_button_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub concrete_combo_Click()
    Dim i As Integer
    i = concrete_combo.ListIndex + 1
    
    name_txt.Caption = concrete(i).name
    description_txt.Caption = concrete(i).description
End Sub

Private Sub Form_Load()
    Call DisableX(frm_concrete_board)
    Call load_database
End Sub


Private Sub load_database()

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
concrete_combo.Clear
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
            concrete_combo.AddItem .name
            If enabler("Concrete", "Database") = .name Then
                concrete_combo.Enabled = False
            End If
            .date = s(1)
            .description = s(2)
            s() = Split(r(1), "#")
            With .wood
                .co2 = s(0)
                .energy = s(1)
                .nox = s(2)
                .so2 = s(3)
                .water = s(4)
            End With
            s() = Split(r(2), "#")
            With .cement
                .co2 = s(0)
                .energy = s(1)
                .nox = s(2)
                .so2 = s(3)
                .water = s(4)
            End With
            s() = Split(r(3), "#")
            With .steel
                .co2 = s(0)
                .energy = s(1)
                .nox = s(2)
                .so2 = s(3)
                .water = s(4)
            End With
            s() = Split(r(4), "#")
            With .agregates
                .co2 = s(0)
                .energy = s(1)
                .nox = s(2)
                .so2 = s(3)
                .water = s(4)
            End With
            s() = Split(r(5), "#")
            With .water
                .co2 = s(0)
                .energy = s(1)
                .nox = s(2)
                .so2 = s(3)
                .water = s(4)
            End With
        End With
    Wend
    
    If num = 1 Then
        num_concrete_entries.Caption = "There's " & CStr(num) & " entry in the concrete database"
    ElseIf num = 0 Then
        num_concrete_entries.Caption = "There aren't any entries in the concrete database. To add entries please select open database in the menu database."
        Next_button.Enabled = False
    Else
        num_concrete_entries.Caption = "There are " & CStr(num) & " entries in the concrete database"
    End If
    concrete_combo.ListIndex = num - 1
Else
    num_concrete_entries.Caption = "There aren't any entries in the concrete database. To add entries please select open database in the menu database."
    concrete_combo.AddItem "No Entrys"
    concrete_combo.ListIndex = 0
    Next_button.Enabled = False
End If
Close #1
End Sub

Private Sub Next_button_Click()
    
    frm_concrete_board2.db_name = concrete_combo.Text
    frm_concrete_board2.db_pos = concrete_combo.ListIndex + 1
    
    Me.Hide
    Unload Me
    frm_concrete_board2.Show 1
End Sub
