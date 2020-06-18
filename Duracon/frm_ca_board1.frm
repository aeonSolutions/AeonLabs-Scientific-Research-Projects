VERSION 5.00
Begin VB.Form frm_ca_board1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DURACON - General Information"
   ClientHeight    =   5790
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7935
   Icon            =   "frm_ca_board1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton sair 
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
      Left            =   4830
      TabIndex        =   22
      Top             =   5220
      Width           =   1425
   End
   Begin VB.Frame TempFrame 
      Caption         =   "Average Anual Temperature"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   4350
      TabIndex        =   21
      Top             =   3600
      Width           =   3435
      Begin VB.OptionButton Testtemp_1 
         Caption         =   "Design Value (21ºC)"
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
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton Testtemp_2 
         Caption         =   "Other  (ºC)  :"
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
         TabIndex        =   11
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox testtemp_val 
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
         Left            =   1650
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame AgeFrame 
      Caption         =   "Age of Structure at assessment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   4380
      TabIndex        =   20
      Top             =   2430
      Width           =   3435
      Begin VB.TextBox testage_val 
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
         Left            =   2070
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Testage_2 
         Caption         =   "Other age (days) :"
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
         Left            =   150
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Project Information"
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7695
      Begin VB.TextBox Description 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   6045
      End
      Begin VB.TextBox Datepj 
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
         Left            =   780
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Prjname 
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
         Left            =   870
         TabIndex        =   1
         Top             =   360
         Width           =   6495
      End
      Begin VB.Label Label3 
         Caption         =   "Date :"
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
         TabIndex        =   19
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "Description :"
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
         TabIndex        =   18
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Name :"
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
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Next_1 
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
      Left            =   6360
      TabIndex        =   12
      Top             =   5220
      Width           =   1425
   End
   Begin VB.Frame TSeries 
      Caption         =   "Time Parameter"
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
      TabIndex        =   14
      Top             =   3630
      Width           =   4095
      Begin VB.ComboBox Timeseries_1 
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
         ItemData        =   "frm_ca_board1.frx":324A
         Left            =   1350
         List            =   "frm_ca_board1.frx":325D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Title_2 
         Caption         =   "Design Life of Structure (years)"
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
         Left            =   690
         TabIndex        =   15
         Top             =   390
         Width           =   3105
      End
   End
   Begin VB.Frame DCoef 
      Caption         =   "Chloride Diffusion Coefficient"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   4095
      Begin VB.OptionButton DCoef_2 
         Caption         =   "Obtained from testing - NT Build 443"
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
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   3735
      End
      Begin VB.OptionButton DCoef_3 
         Caption         =   "Obtained from chloride profile"
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000015&
      X1              =   210
      X2              =   7470
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Label Title_1 
      Caption         =   "   MODEL PARAMETERS"
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
      Left            =   2790
      TabIndex        =   0
      Top             =   1890
      Width           =   2805
   End
End
Attribute VB_Name = "frm_ca_board1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim doc As Integer

Private Sub Datepj_Change()
Dim doc As Integer

doc = current_form
On Error GoTo DiskErrorHandler
With doc_props(doc)
    If Datepj.Text <> "" Then
        .datepjt = Datepj.Text
    End If
End With
Exit Sub
    
DiskErrorHandler:
    Dim m$
    Dim WhatToDo%
    Beep
    Select Case Err.Number
        Case 13
            m$ = "Numeric (integer) values only!"
        Case 6
            m$ = "Numeric value too large (10 digits allowed)"
        Case Else
            m$ = "Please report this bug: " & "Error " & Err.Number & " "
    End Select
    
    WhatToDo% = MsgBox(m$, vbOK + vbCritical, "Duracon")
    If WhatToDo% = vbOKOnly Then
        Datepj.Text = Empty
    Else
        Datepj.Text = Empty
        Exit Sub
    End If
End Sub




Private Sub Description_Change()
Dim doc As Integer

doc = current_form
If Description.Text <> "" Then
    doc_props(doc).descrip = Description.Text
End If
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()

Dim i As Integer
Dim doc As Integer
doc = current_form

last_window = "frm_ca_board1"
Call DisableX(frm_ca_board1)
With doc_props(doc)
    .iseedv = 1
    .nsimul = 10000
    .seed = 0
    ' Desgin value default values
    .idifcoef = 1
    'age of structure during assessment
    .prmvone(6) = 28
    .prmvtwo(6) = 0
    'design life of structure - default is 10 years
    .tseriev = 10
    .prmvone(5) = 0
    If .frm_ca_board1_values.values And .frm_ca_board1_values.project_name <> "" Then ' there's already input data stored
        With .frm_ca_board1_values
            Prjname.Text = .project_name
            Description.Text = .Description
            Datepj.Text = .project_date
            
            If .testage_val = "N/A" Then
                Testage_2.Value = False
            Else
                testage_val.Text = .testage_val
                Testage_2.Value = True
            End If
            
            If .testtemp_val = "N/A" Then
                Testtemp_1.Value = True
                Testtemp_2.Value = False
            Else
                Testtemp_2.Value = True
                Testtemp_1.Value = False
                testtemp_val.Text = .testtemp_val
            End If
            
            Timeseries_1.ListIndex = .Timeseries_1
            If .cdc = 2 Then
                DCoef_2.Value = True
            ElseIf .cdc = 3 Then
                DCoef_3.Value = True
            End If
        End With
    Else
        Timeseries_1.ListIndex = 0
    End If
    .kk = 0
    If .datepjt = "" Then
        .datepjt = Date
        Datepj.Text = Date
    Else
        Datepj.Text = .datepjt
    End If
    
    If .tt = 99 Then
        If .idifcoef = 0 Then
            If .prmvone(6) <> 28 Then
                DCoef_2.Value = True
            End If
        End If 'issedv
        If .idifcoef = 2 Then
            DCoef_3.Value = True
        End If 'issedv
    
        If ((.prmvone(6) <> 28) And (.prmvone(6) <> 63)) Then
            Testage_2.Value = True
            testage_val.Text = .prmvone(6)
        End If
    
        If .prmvone(7) <> 21 Then
            Testtemp_2.Value = True
            testtemp_val.Text = .prmvone(7)
        End If
    
    End If
    
    If DCoef_3.Value = False Then
        .prfpred = 0
        .prfsz = 0
        .kk = 123
    End If
End With
End Sub

Private Sub Next_1_Click()
Dim doc As Integer


doc = current_form
If save_data Then
    doc_props(doc).frm_ca_board1_values.values = True
    FState(doc).saved = False
Else
    Exit Sub
End If
With doc_props(doc)
    If Prjname.Text = "" Or Datepj.Text = "yyyymmdd" Or Timeseries_1.Text = "Choose" Then
        MsgBox ("Fill in all cells!"), vbOK + vbCritical, "Duracon"
       .frm_ca_board1_values.ready = False
    Else
       .frm_ca_board1_values.ready = True
        frm_ca_board1.Hide
        If DCoef_3.Value = True Then
            frm_ca_board2.Show 1
            Unload Me
        Else
            frm_ca_board3.Show 1
            Unload Me
        End If
    End If
End With

End Sub

Private Sub Prjname_Change()
If Prjname.Text <> "" Then
    With doc_props(current_form)
        .nprojt = Prjname.Text
    End With
End If
End Sub
Private Function save_data() As Boolean
doc = current_form
FState(doc).values = True
Dim tmp() As String

save_data = True
With doc_props(doc)
    If testtemp_val.Text <> "" Then
        tmp() = Split(testtemp_val.Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            testtemp_val.Text = "0" & testtemp_val.Text
        End If
        tmp() = Split(testtemp_val.Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            testtemp_val.Text = "0" & testtemp_val.Text
        End If
        If IsNumeric(testtemp_val.Text) Then
            On Error Resume Next
            .prmvone(7) = CSng(testtemp_val.Text)
            .frm_ca_board1_values.testtemp_val = testtemp_val.Text
            Testtemp_2.Value = True
            Testtemp_1.Value = False
            If Err.Number <> 0 Then
                If Err.Number = 6 Then
                    MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                    testtemp_val.SetFocus
                    save_data = False
                    Exit Function
                Else
                    MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOK + vbCritical, "Duracon"
                    testtemp_val.SetFocus
                    save_data = False
                    Exit Function
                End If
            End If
        End If
    Else
        If Testtemp_2.Value = True Then
            MsgBox "Only Positive numeric Values allowed!", vbOK + vbCritical, "Duracon"
            testtemp_val.SetFocus
            save_data = False
            Exit Function
        End If
    End If
    If testage_val.Text <> "" Then
        tmp() = Split(testage_val.Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            testage_val.Text = "0" & testage_val.Text
        End If
        tmp() = Split(testage_val.Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            testage_val.Text = "0" & testage_val.Text
        End If
        If IsNumeric(testage_val.Text) Then
            On Error Resume Next
            doc_props(doc).prmvone(6) = CSng(testage_val.Text)
            doc_props(doc).frm_ca_board1_values.testage_val = testtemp_val.Text
            Testage_2.Value = True
            If Err.Number <> 0 Then
                If Err.Number = 6 Then
                    MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                    testage_val.SetFocus
                    save_data = False
                    Exit Function
                Else
                    MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOK + vbCritical, "Duracon"
                    testage_val.SetFocus
                    save_data = False
                    Exit Function
                End If
            End If
        End If
    Else
       If Testage_2.Value = True Then
            MsgBox "Only Positive numeric Values allowed!", vbOK + vbCritical, "Duracon"
            testage_val.SetFocus
            save_data = False
            Exit Function
        End If
    End If
    .frm_ca_board1_values.project_name = Prjname.Text
    .frm_ca_board1_values.Description = Description.Text
    .frm_ca_board1_values.project_date = Datepj.Text
    
    
    If DCoef_2.Value = True Then
        If .frm_ca_board2_values.values Then
                .frm_ca_board2_values.values = False
        End If
        .frm_ca_board1_values.cdc = 2
        .idifcoef = 0
        .kk = 123
        If .tt <> 99 Then
            .prmvone(6) = 63 ' 28 day maturity + at least 35 day ponding
            .prmvtwo(6) = 0
        End If
    End If
    
    If DCoef_3.Value = True Then
        .frm_ca_board1_values.cdc = 3
        .idifcoef = 2
        If .tt <> 99 Then
            .prmvone(6) = 28
            .prmvtwo(6) = 0
        End If
    End If

    If testage_val.Text <> "" Then
        .frm_ca_board1_values.testage_val = testage_val.Text
        .prmvone(6) = CSng(testage_val.Text)
    Else
        .frm_ca_board1_values.testage_val = "N/A"
    End If
    
    .iseedv = 1
    .nsimul = 10000
    
    If Testtemp_1.Value = True Then
        .frm_ca_board1_values.testtemp_val = "N/A"
        .prmvone(7) = 21
        .prmvtwo(7) = 0
    Else
        .frm_ca_board1_values.testtemp_val = testtemp_val.Text
    End If
End With

End Function
Private Sub sair_Click()
    doc = current_form
    If save_data Then
        doc_props(doc).frm_ca_board1_values.values = True
    Else
        Exit Sub
    End If
    Call refresh_lista(doc)
    FState(doc).values = False
    Unload Me
End Sub



Private Sub testage_val_Change()
Dim tmp() As String
Dim tmp2 As String

With doc_props(doc)
    If testage_val.Text <> "" Then
        tmp2 = testage_val.Text
        tmp() = Split(testage_val.Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            tmp2 = "0" & testage_val.Text
        End If
        tmp() = Split(testage_val.Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            tmp2 = "0" & testage_val.Text
        End If
        If IsNumeric(tmp2) Then
            On Error Resume Next
            doc_props(doc).prmvone(6) = CSng(tmp2)
            doc_props(doc).frm_ca_board1_values.testage_val = tmp2
            If Err.Number <> 0 Then
                Exit Sub
            End If
            Testage_2.Value = True
        Else
            Testage_2.Value = False
        End If
    End If
End With
End Sub

Private Sub Testtemp_1_Click()
Dim doc As Integer
With doc_props(current_form)
    If Testtemp_1.Value = True Then
         .prmvone(7) = 21
         .prmvtwo(7) = 0
         .frm_ca_board1_values.testtemp_val = "N/A"
         Testtemp_2.Value = False
         testtemp_val.Text = ""
    End If
End With
End Sub

Private Sub testtemp_val_Change()
Dim tmp() As String
Dim tmp2 As String

With doc_props(doc)
    If testtemp_val.Text <> "" Then
        tmp2 = testtemp_val.Text
        tmp() = Split(testtemp_val.Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            tmp2 = "0" & testtemp_val.Text
        End If
        tmp() = Split(testtemp_val.Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            tmp2 = "0" & testtemp_val.Text
        End If
        If IsNumeric(tmp2) Then
            On Error Resume Next
            .prmvone(7) = CSng(tmp2)
            .frm_ca_board1_values.testtemp_val = tmp2
            If Err.Number <> 0 Then
            
             Exit Sub
            End If
            Testtemp_2.Value = True
            Testtemp_1.Value = False
        Else
            Testtemp_1.Value = True
            Testtemp_2.Value = False
        End If
    End If
End With

End Sub


Private Sub Timeseries_1_Click()
Dim doc As Integer

With doc_props(current_form)
    .tseriev = 10
    .prmvone(5) = 0
    .prmvtwo(5) = 50
    .frm_ca_board1_values.Timeseries_1 = 0

    If Timeseries_1.Text <> "" Then
        .prmvtwo(5) = Val(Timeseries_1.Text)
        .frm_ca_board1_values.Timeseries_1 = Timeseries_1.ListIndex
        .tseriev = (.prmvtwo(5) / 5)
    End If
End With

End Sub

