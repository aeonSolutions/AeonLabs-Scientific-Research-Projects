VERSION 5.00
Begin VB.Form frm_ca_board3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DURACON - Distribution Data"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ca_board3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton back_button 
      Caption         =   "лл Back"
      Height          =   375
      Left            =   4890
      TabIndex        =   16
      Top             =   5520
      Width           =   1425
   End
   Begin VB.CommandButton close_button 
      Caption         =   "Finish"
      Height          =   375
      Left            =   6420
      TabIndex        =   17
      Top             =   5520
      Width           =   1425
   End
   Begin VB.Frame Title_2 
      Caption         =   "Distribution Data Variables"
      Height          =   4725
      Left            =   120
      TabIndex        =   18
      Top             =   510
      Width           =   7695
      Begin VB.ComboBox Distype 
         Height          =   315
         Index           =   4
         ItemData        =   "frm_ca_board3.frx":324A
         Left            =   420
         List            =   "frm_ca_board3.frx":325A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3360
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
         Height          =   315
         Index           =   3
         ItemData        =   "frm_ca_board3.frx":3286
         Left            =   420
         List            =   "frm_ca_board3.frx":3296
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1350
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
         Height          =   315
         Index           =   2
         ItemData        =   "frm_ca_board3.frx":32C2
         Left            =   420
         List            =   "frm_ca_board3.frx":32D2
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2700
         Width           =   1695
      End
      Begin VB.TextBox Param1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   14
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Param1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3960
         TabIndex        =   11
         Top             =   1350
         Width           =   1695
      End
      Begin VB.TextBox Param1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   8
         Top             =   2700
         Width           =   1695
      End
      Begin VB.TextBox Param1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   5
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Param2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   5850
         TabIndex        =   15
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Param2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   5850
         TabIndex        =   12
         Top             =   1350
         Width           =   1695
      End
      Begin VB.TextBox Param2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   5850
         TabIndex        =   9
         Top             =   2700
         Width           =   1695
      End
      Begin VB.TextBox Param2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   5850
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
         Height          =   315
         Index           =   1
         ItemData        =   "frm_ca_board3.frx":32FE
         Left            =   420
         List            =   "frm_ca_board3.frx":330E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox Param1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   2
         Top             =   4050
         Width           =   1695
      End
      Begin VB.TextBox Param2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   5850
         TabIndex        =   3
         Top             =   4050
         Width           =   1695
      End
      Begin VB.ComboBox Distype 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frm_ca_board3.frx":333A
         Left            =   420
         List            =   "frm_ca_board3.frx":334A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4050
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Xc (mm)"
         Height          =   255
         Left            =   2490
         TabIndex        =   33
         Top             =   4050
         Width           =   885
      End
      Begin VB.Label Label21 
         Caption         =   "a"
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
         Left            =   2610
         TabIndex        =   32
         Top             =   3360
         Width           =   135
      End
      Begin VB.Label Label19 
         Caption         =   "(-)"
         Height          =   255
         Left            =   2760
         TabIndex        =   31
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "Diffusion Coefficient"
         Height          =   255
         Left            =   210
         TabIndex        =   30
         Top             =   1800
         Width           =   2265
      End
      Begin VB.Label Label12 
         Caption         =   "Age Effect Diffusion"
         Height          =   285
         Left            =   210
         TabIndex        =   29
         Top             =   3120
         Width           =   2235
      End
      Begin VB.Label Label11 
         Caption         =   "Cs (% wt./cem)"
         Height          =   255
         Left            =   2340
         TabIndex        =   28
         Top             =   1350
         Width           =   1725
      End
      Begin VB.Label Label10 
         Caption         =   "Surface Chloride concentration"
         Height          =   285
         Left            =   210
         TabIndex        =   27
         Top             =   1110
         Width           =   2805
      End
      Begin VB.Label Label9 
         Caption         =   "Deff (e-12 m2/s)"
         Height          =   255
         Left            =   2250
         TabIndex        =   26
         Top             =   2040
         Width           =   1635
      End
      Begin VB.Label Label8 
         Caption         =   "Critical Chloride Concentration"
         Height          =   255
         Left            =   210
         TabIndex        =   25
         Top             =   2460
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "Ccr (% wt./cem)"
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   2700
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "Concrete Cover"
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   3810
         Width           =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000015&
         X1              =   120
         X2              =   7560
         Y1              =   750
         Y2              =   750
      End
      Begin VB.Label Label5 
         Caption         =   "Variables"
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Parameter 1"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   570
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Parameters 2"
         Height          =   255
         Left            =   6060
         TabIndex        =   20
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Distribution type"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   570
         Width           =   1635
      End
   End
   Begin VB.Label Title_1 
      Caption         =   "   DISTRIBUTION DATA"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frm_ca_board3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private doc As Integer



Private Sub back_button_Click()
Dim i As Integer

doc_props(doc).frm_ca_board3_values.values = True
Call save_data(True)
doc_props(doc).frm_ca_board3_values.ready = False

If last_window = "frm_ca_board1" Then
    frm_ca_board3.Hide
    frm_ca_board1.Show 1
    Unload Me
ElseIf last_window = "frm_ca_board2" Then
    frm_ca_board3.Hide
    frm_ca_board2.Show 1
    Unload Me
End If
End Sub

Private Function save_data(silent As Boolean) As Boolean
Dim i As Integer
Dim tmp() As String

save_data = True
For i = 0 To 4
With doc_props(doc)
    If Param1(i).Text <> "" Then
        tmp() = Split(Param1(i).Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            Param1(i).Text = "0" & Param1(i).Text
        End If
        tmp() = Split(Param1(i).Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            Param1(i).Text = "0" & Param1(i).Text
        End If
        If IsNumeric(Param1(i).Text) Then
            On Error Resume Next
            .prmvone(i) = Val(Param1(i).Text)
            If Err.Number <> 0 Then
                If Err.Number = 6 Then
                    If Not silent Then MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                    Param1(i).SetFocus
                    save_data = False
                    Exit Function
                Else
                    If Not silent Then MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOK + vbCritical, "Duracon"
                    Param1(i).SetFocus
                    save_data = False
                    Exit Function
                End If
            End If
        End If
    Else
        If Not silent Then MsgBox "Only Positive numeruc values allowed", vbOK + vbCritical, "Duracon"
        Param1(i).SetFocus
        save_data = False
        Exit Function
    End If
    If Param1(i).Text <> "" Then
        tmp() = Split(Param2(i).Text, ".")
        If tmp(0) = "." Or tmp(0) = "" Then
            Param2(i).Text = "0" & Param2(i).Text
        End If
        tmp() = Split(Param2(i).Text, ",")
        If tmp(0) = "," Or tmp(0) = "" Then
            Param2(i).Text = "0" & Param2(i).Text
        End If
        If IsNumeric(Param2(i).Text) Then
            On Error Resume Next
            .prmvtwo(i) = Val(Param2(i).Text)
            If Err.Number <> 0 Then
                If Err.Number = 6 Then
                    If Not silent Then MsgBox "Numeric value too large", vbOK + vbCritical, "Duracon"
                    Param2(i).SetFocus
                    save_data = False
                    Exit Function
                Else
                    If Not silent Then MsgBox "Please report this bug: " & "Error " & Err.Number & " ", vbOK + vbCritical, "Duracon"
                    Param2(i).SetFocus
                    save_data = False
                    Exit Function
                End If
            End If
        End If
    Else
        If Not silent Then MsgBox "Only Positive numeruc values allowed", vbOK + vbCritical, "Duracon"
        Param2(i).SetFocus
        save_data = False
        Exit Function
    End If

End With
Next i
With doc_props(doc)
    .frm_ca_board3_values.values = True
    For i = 0 To 4
        If Distype(i).Text = "Deterministic" Then .prmdistn(i) = 7
        If Distype(i).Text = "Normal" Then .prmdistn(i) = 0
        If Distype(i).Text = "Lognormal" Then .prmdistn(i) = 1
        If Distype(i).Text = "Beta" Then .prmdistn(i) = 2
        .frm_ca_board3_values.Distype(i) = Distype(i).ListIndex
    Next i
    
        .prmdistn(5) = 7
        .prmdistn(6) = 7
        .prmdistn(7) = 0
End With
End Function
Private Sub close_button_Click()
Dim i As Integer

If save_data(True) Then
    doc_props(doc).frm_ca_board3_values.values = True
    doc_props(doc).frm_ca_board3_values.ready = False
Else
    doc_props(doc).frm_ca_board3_values.values = False
    doc_props(doc).frm_ca_board3_values.ready = False
End If
On Error GoTo DiskErrorHandler
With doc_props(doc)
    If .tseriev = 1 And .prmvone(5) = 0 Then
        MsgBox ("Time (parameter 1) igual to zero!"), vbOK + vbCritical, "Duracon"
    Exit Sub
    End If
    
    If .tseriev = 1 And .prmvtwo(5) <> 0 Then
        MsgBox ("Time (parameter 2) different than zero!"), vbOK + vbCritical, "Duracon"
        Exit Sub
    End If
    
    If .tseriev > 1 And .prmvtwo(5) = 0 Then
        MsgBox ("Time (parameter 2) igual to zero!"), vbOK + vbCritical, "Duracon"
    Exit Sub
    End If
    
    If .tseriev > 1 And (.prmvtwo(5) <= .prmvone(5)) Then
        MsgBox ("Time (parameter 2) smaller than (parameter 1)!"), vbOK + vbCritical, "Duracon"
    Exit Sub
    End If
    
    .tt = 1
    
End With
doc_props(doc).frm_ca_board3_values.ready = True
Call refresh_lista(doc)
FState(doc).values = True
Unload Me
Exit Sub

DiskErrorHandler:
Dim m$
Dim WhatToDo%
Beep
Select Case Err.Number
    Case 13
        m$ = "All parameters must be filled in.Are You sure you want to close?"
        WhatToDo% = MsgBox(m$, vbOK + vbCritical, "Duracon")
        If WhatToDo% = vbYes Then
            doc_props(doc).frm_ca_board3_values.ready = False
            doc_props(doc).frm_ca_board3_values.values = False
            Unload Me
        End If
    Case Else
        m$ = "Please report this bug: " & "Error " & Err.Number & " "
        doc_props(doc).frm_ca_board3_values.ready = False
        WhatToDo% = MsgBox(m$, vbOK + vbCritical, "Duracon")
End Select

End Sub

Private Sub Distype_Click(Index As Integer)
With doc_props(doc)
    If Distype(Index).Text <> "" Then
        If Distype(Index).Text = "Deterministic" Then .prmdistn(Index) = 7
        If Distype(Index).Text = "Normal" Then .prmdistn(Index) = 0
        If Distype(Index).Text = "Lognormal" Then .prmdistn(Index) = 1
        If Distype(Index).Text = "Beta" Then .prmdistn(Index) = 2
    End If
End With
End Sub


Private Sub Form_Load()
Dim i As Integer

doc = current_form
Call DisableX(frm_ca_board3)

With doc_props(doc)
    If .frm_ca_board3_values.values Then
        For i = 0 To 4
            Distype(i).ListIndex = .frm_ca_board3_values.Distype(i)
            Param1(i).Text = CStr(.prmvone(i))
            Param2(i).Text = CStr(.prmvtwo(i))
        Next i
    Else
        For i = 0 To 4
            Distype(i).ListIndex = 1
        Next i
    End If
End With
With doc_props(doc)
If .frm_ca_board1_values.cdc = 3 Then
    Distype(1).Enabled = False
    Distype(3).Enabled = False
    Param1(1).Enabled = False
    Param1(3).Enabled = False
    Param2(1).Enabled = False
    Param2(3).Enabled = False
End If
End With
End Sub

