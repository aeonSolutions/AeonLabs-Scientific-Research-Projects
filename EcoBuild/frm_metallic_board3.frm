VERSION 5.00
Begin VB.Form frm_metallic_board3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pillars"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6585
   Icon            =   "frm_metallic_board3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   345
      Left            =   2550
      TabIndex        =   12
      Top             =   3360
      Width           =   1125
   End
   Begin VB.CommandButton close 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3360
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   3870
      TabIndex        =   8
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pillars"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   300
      TabIndex        =   3
      Top             =   420
      Width           =   6015
      Begin VB.TextBox weight_txt 
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
         Left            =   4770
         TabIndex        =   1
         Top             =   660
         Width           =   1005
      End
      Begin VB.ComboBox profile_combo 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   660
         Width           =   2145
      End
      Begin VB.TextBox num_beams 
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
         Left            =   4770
         TabIndex        =   4
         Top             =   1890
         Width           =   1005
      End
      Begin VB.TextBox lenght_txt 
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
         Left            =   1890
         TabIndex        =   2
         Top             =   1860
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Weight (Kg/ml) :"
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
         Left            =   3300
         TabIndex        =   11
         Top             =   690
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Type:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Number of Pillars :"
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
         Left            =   3090
         TabIndex        =   7
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Label Label22 
         Caption         =   "Pillar Lenght (m) :"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1890
         Width           =   1665
      End
      Begin VB.Label Label19 
         Caption         =   "Cross-Section :"
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
         Left            =   330
         TabIndex        =   5
         Top             =   330
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frm_metallic_board3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private profiles(1 To 133, 1 To 2) As String

Private Sub Close_Click()
    frm_metallic_board3.Hide
    Unload Me
End Sub


Private Sub Command1_Click()
Dim doc As Integer

If Not validate_fields() Then
    Exit Sub
End If
doc = current_form()
FState(doc).count = FState(doc).count + 1
With document(doc)
    With .lista
        .Row = FState(doc).count
        .Col = 0
        .Text = "Met.Pillar"
        .Col = 1
        .Text = num_beams.Text
        .Col = 2
        .Text = weight_txt.Text
        .Col = 3
        .Text = lenght_txt.Text
    End With
End With
End Sub

Private Sub Command2_Click()
Unload Me
frm_metallic_board2.Show 1
End Sub

Private Sub Form_Load()
 Dim i As Integer
 
Call DisableX(frm_metallic_board3)
Call load_profiles
profile_combo.AddItem "Other Profile"
For i = 1 To 133
 profile_combo.AddItem profiles(i, 1)
Next i
profile_combo.ListIndex = 0
End Sub

Private Function validate_fields() As Boolean

validate_fields = True
If Not IsNumeric(lenght_txt.Text) Then
    validate_fields = False
    lenght_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(weight_txt.Text) Then
    validate_fields = False
    weight_txt.SetFocus
    Exit Function
End If
If Not IsNumeric(num_beams.Text) Then
    validate_fields = False
    num_beams.SetFocus
    Exit Function
End If
End Function

Private Sub load_profiles()
'INP profiles: Designation / weight (kg/m)
profiles(1, 1) = "INP 80"
profiles(1, 2) = "5,95"
profiles(2, 1) = "INP 100"
profiles(2, 2) = "8,32"
profiles(3, 1) = "INP 120"
profiles(3, 2) = "11,2"
profiles(4, 1) = "INP 140"
profiles(4, 2) = "14,4"
profiles(5, 1) = "INP 160"
profiles(5, 2) = "17,9"
profiles(6, 1) = "INP 180"
profiles(6, 2) = "21,9"
profiles(7, 1) = "INP 200"
profiles(7, 2) = "26,3"
profiles(8, 1) = "INP 220"
profiles(8, 2) = "31,1"
profiles(9, 1) = "INP 240"
profiles(9, 2) = "36,2"
profiles(10, 1) = "INP 260"
profiles(10, 2) = "41,9"
profiles(11, 1) = "INP 280"
profiles(11, 2) = "48,0"
profiles(12, 1) = "INP 300"
profiles(12, 2) = "54,2"
profiles(13, 1) = "INP 320"
profiles(13, 2) = "61,1"
profiles(14, 1) = "INP 340"
profiles(14, 2) = "68,1"
profiles(15, 1) = "INP 360"
profiles(15, 2) = "76,2"
profiles(16, 1) = "INP 400"
profiles(16, 2) = "92,6"
profiles(17, 1) = "INP 450"
profiles(17, 2) = "115"
profiles(18, 1) = "INP 500"
profiles(18, 2) = "141"
profiles(19, 1) = "INP 550"
profiles(19, 2) = "167"
profiles(20, 1) = "INP 600"
profiles(20, 2) = "199"

'IPE profiles: Designation / weight (kg/m)
profiles(21, 1) = "IPE 80"
profiles(21, 2) = "6,0"
profiles(22, 1) = "IPE 100"
profiles(22, 2) = "8,1"
profiles(23, 1) = "IPE 120"
profiles(23, 2) = "10,4"
profiles(24, 1) = "IPE 140"
profiles(24, 2) = "12,9"
profiles(25, 1) = "IPE 160"
profiles(25, 2) = "15,6"
profiles(26, 1) = "IPE 180"
profiles(26, 2) = "18,8"
profiles(27, 1) = "IPE 200"
profiles(27, 2) = "22,4"
profiles(28, 1) = "IPE 220"
profiles(28, 2) = "26,2"
profiles(29, 1) = "IPE 240"
profiles(29, 2) = "30,7"
profiles(30, 1) = "IPE 270"
profiles(30, 2) = "36,1"
profiles(31, 1) = "IPE 300"
profiles(31, 2) = "42,2"
profiles(32, 1) = "IPE 330"
profiles(32, 2) = "49,1"
profiles(33, 1) = "IPE 360"
profiles(33, 2) = "57,1"
profiles(34, 1) = "IPE 400"
profiles(34, 2) = "66,3"
profiles(35, 1) = "IPE 450"
profiles(35, 2) = "77,6"
profiles(36, 1) = "IPE 500"
profiles(36, 2) = "80,7"
profiles(37, 1) = "IPE 550"
profiles(37, 2) = "106"
profiles(38, 1) = "IPE 600"
profiles(38, 2) = "122"

'HE profiles: Designation / weight (kg/m)
profiles(39, 1) = "HE 100 A"
profiles(39, 2) = "16,7"
profiles(40, 1) = "HE 100 B"
profiles(40, 2) = "20,4"
profiles(41, 1) = "HE 100 M"
profiles(41, 2) = "41,8"
profiles(42, 1) = "HE 120 A"
profiles(42, 2) = "19,9"
profiles(43, 1) = "HE 120 B"
profiles(43, 2) = "26,7"
profiles(44, 1) = "HE 120 M"
profiles(44, 2) = "52,1"
profiles(45, 1) = "HE 140 A"
profiles(45, 2) = "24,7"
profiles(46, 1) = "HE 140 B"
profiles(46, 2) = "33,7"
profiles(47, 1) = "HE 140 M"
profiles(47, 2) = "63,2"
profiles(48, 1) = "HE 160 A"
profiles(48, 2) = "30,4"
profiles(49, 1) = "HE 160 B"
profiles(49, 2) = "42,6"
profiles(50, 1) = "HE 160 M"
profiles(50, 2) = "76,2"
profiles(51, 1) = "HE 180 A"
profiles(51, 2) = "35,5"
profiles(52, 1) = "HE 180 B"
profiles(52, 2) = "51,2"
profiles(53, 1) = "HE 180 M"
profiles(53, 2) = "88,9"
profiles(54, 1) = "HE 200 A"
profiles(54, 2) = "42,3"
profiles(55, 1) = "HE 200 B"
profiles(55, 2) = "61,3"
profiles(56, 1) = "HE 200 M"
profiles(56, 2) = "103"
profiles(57, 1) = "HE 220 A"
profiles(57, 2) = "50,5"
profiles(58, 1) = "HE 220 B"
profiles(58, 2) = "71,5"
profiles(59, 1) = "HE 220 M"
profiles(59, 2) = "117"
profiles(60, 1) = "HE 240 A"
profiles(60, 2) = "60,3"
profiles(61, 1) = "HE 240 B"
profiles(61, 2) = "83,2"
profiles(62, 1) = "HE 240 M"
profiles(62, 2) = "157"
profiles(63, 1) = "HE 260 A"
profiles(63, 2) = "68,2"
profiles(64, 1) = "HE 260 B"
profiles(64, 2) = "93,0"
profiles(65, 1) = "HE 260 M"
profiles(65, 2) = "172"
profiles(66, 1) = "HE 280 A"
profiles(66, 2) = "76,4"
profiles(67, 1) = "HE 280 B"
profiles(67, 2) = "103"
profiles(68, 1) = "HE 280 M"
profiles(68, 2) = "189"
profiles(69, 1) = "HE 300 A"
profiles(69, 2) = "88,3"
profiles(70, 1) = "HE 300 B"
profiles(70, 2) = "117"
profiles(71, 1) = "HE 300 C"
profiles(71, 2) = "177"
profiles(72, 1) = "HE 300 M"
profiles(72, 2) = "238"

profiles(73, 1) = "HE 320 A"
profiles(73, 2) = "97,6"
profiles(74, 1) = "HE 320 B"
profiles(74, 2) = "127"
profiles(75, 1) = "HE 320 M"
profiles(75, 2) = "245"
profiles(76, 1) = "HE 340 A"
profiles(76, 2) = "105"
profiles(77, 1) = "HE 340 B"
profiles(77, 2) = "134"
profiles(78, 1) = "HE 340 M"
profiles(78, 2) = "248"
profiles(79, 1) = "HE 360 A"
profiles(79, 2) = "112"
profiles(80, 1) = "HE 360 B"
profiles(80, 2) = "142"
profiles(81, 1) = "HE 360 M"
profiles(81, 2) = "250"
profiles(82, 1) = "HE 400 A"
profiles(82, 2) = "125"
profiles(83, 1) = "HE 400 B"
profiles(83, 2) = "155"
profiles(84, 1) = "HE 400 M"
profiles(84, 2) = "256"
profiles(85, 1) = "HE 450 A"
profiles(85, 2) = "140"
profiles(86, 1) = "HE 450 B"
profiles(86, 2) = "171"
profiles(87, 1) = "HE 450 M"
profiles(87, 2) = "263"
profiles(88, 1) = "HE 500 A"
profiles(88, 2) = "155"
profiles(89, 1) = "HE 500 B"
profiles(89, 2) = "187"
profiles(90, 1) = "HE 500 M"
profiles(90, 2) = "270"
profiles(91, 1) = "HE 550 A"
profiles(91, 2) = "166"
profiles(92, 1) = "HE 550 B"
profiles(92, 2) = "199"
profiles(93, 1) = "HE 550 M"
profiles(93, 2) = "278"
profiles(94, 1) = "HE 600 A"
profiles(94, 2) = "178"
profiles(95, 1) = "HE 600 B"
profiles(95, 2) = "212"
profiles(96, 1) = "HE 600 M"
profiles(96, 2) = "285"
profiles(97, 1) = "HE 650 A"
profiles(97, 2) = "190"
profiles(98, 1) = "HE 650 B"
profiles(98, 2) = "225"
profiles(99, 1) = "HE 650 M"
profiles(99, 2) = "293"
profiles(100, 1) = "HE 700 A"
profiles(100, 2) = "204"
profiles(101, 1) = "HE 700 B"
profiles(101, 2) = "241"
profiles(102, 1) = "HE 700 M"
profiles(102, 2) = "301"
profiles(103, 1) = "HE 800 A"
profiles(103, 2) = "224"
profiles(104, 1) = "HE 800 B"
profiles(104, 2) = "262"
profiles(105, 1) = "HE 800 M"
profiles(105, 2) = "317"
profiles(106, 1) = "HE 900 A"
profiles(106, 2) = "252"
profiles(107, 1) = "HE 900 B"
profiles(107, 2) = "291"
profiles(108, 1) = "HE 900 M"
profiles(108, 2) = "333"
profiles(109, 1) = "HE 1000 A"
profiles(109, 2) = "272"
profiles(110, 1) = "HE 1000 B"
profiles(110, 2) = "314"
profiles(111, 1) = "HE 1000 M"
profiles(111, 2) = "349"

'UNP profiles: Designation / weight (kg/m)
profiles(112, 1) = "UNP 30a"
profiles(112, 2) = "1,74"
profiles(113, 1) = "UNP 30"
profiles(113, 2) = "4,27"
profiles(114, 1) = "UNP 40a"
profiles(114, 2) = "2,87"
profiles(115, 1) = "UNP 40"
profiles(115, 2) = "4,87"
profiles(116, 1) = "UNP 50a"
profiles(116, 2) = "3,86"
profiles(117, 1) = "UNP 50"
profiles(117, 2) = "5,59"
profiles(118, 1) = "UNP 60a"
profiles(118, 2) = "5,07"
profiles(119, 1) = "UNP 65"
profiles(119, 2) = "7,09"
profiles(120, 1) = "UNP 80"
profiles(120, 2) = "8,64"
profiles(121, 1) = "UNP 100"
profiles(121, 2) = "10,6"
profiles(122, 1) = "UNP 120"
profiles(122, 2) = "13,4"
profiles(123, 1) = "UNP 140"
profiles(123, 2) = "16,0"
profiles(124, 1) = "UNP 160"
profiles(124, 2) = "18,8"
profiles(125, 1) = "UNP 180"
profiles(125, 2) = "22,0"
profiles(126, 1) = "UNP 200"
profiles(126, 2) = "25,3"
profiles(127, 1) = "UNP 220"
profiles(127, 2) = "29,4"
profiles(128, 1) = "UNP 240"
profiles(128, 2) = "33,2"
profiles(129, 1) = "UNP 260"
profiles(129, 2) = "37,9"
profiles(130, 1) = "UNP 280"
profiles(130, 2) = "41,8"
profiles(131, 1) = "UNP 300"
profiles(131, 2) = "46,2"
profiles(132, 1) = "UNP 350"
profiles(132, 2) = "60,6"
profiles(133, 1) = "UNP 400"
profiles(133, 2) = "71,8"
 End Sub

Private Sub profile_combo_Click()
    On Error Resume Next
    If profile_combo.ListIndex = 0 Then
        weight_txt.Text = ""
    Else
        weight_txt.Text = profiles(profile_combo.ListIndex, 2)
    End If
End Sub
