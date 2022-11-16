VERSION 5.00
Begin VB.Form delete_material 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete material"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   Icon            =   "delete_material.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3015
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3990
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Delete material"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.TextBox line_txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label info_txt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Line to remove"
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
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
   End
End
Attribute VB_Name = "delete_material"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOK_Click()
Dim line As Integer
Dim arraycount As Integer
Dim i As Integer

If Not IsNumeric(line_txt) Then
   line_txt.SetFocus
   Exit Sub
End If
line = Val(line_txt)
' Cycle through the document array
arraycount = UBound(document)
For i = 1 To arraycount
     If FState(i).Dirty Then
        Exit For
     End If
Next
FState(i).calculated = False
If FState(i).Conta = 1 Then
    Exit Sub
End If

If FState(i).Conta - 1 < i Then
    line_txt.SetFocus
    Exit Sub
End If
FState(i).Conta = FState(i).Conta - 1
With document(i)
   .lista.RemoveItem (line)
   .lista.AddItem ""
   .lista.col = 0
   If FState(i).Conta > 1 Then
     For j = 1 To FState(i).Conta - 1
       .lista.CellAlignment = 4
       .lista.CellFontBold = True
       .lista.row = j
       .lista.Text = Str(j)
     Next j
   End If
   .lista.Refresh
End With

Unload Me
End Sub

Private Sub form_load()
  Dim i As Integer
    
  i = current_form()
  Call DisableX(delete_material)

    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    delete_material.Caption = delete_material.Caption & " - " & document(i).Caption
    If FState(i).Conta = 1 Then
       info_txt.Caption = "There are no material to remove in the current document. Please select the Cancel Button"
    Else
       info_txt.Caption = "The valid line numbers to remove in the current document are form 1 to " & Str(FState(i).Conta - 1)
    End If
End Sub

