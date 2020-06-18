VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChild 
   Caption         =   "frmChild"
   ClientHeight    =   9000
   ClientLeft      =   1515
   ClientTop       =   1575
   ClientWidth     =   10410
   Icon            =   "frmChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   540
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dialogs 
      Left            =   300
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
      Height          =   10605
      Left            =   30
      ScaleHeight     =   10605
      ScaleWidth      =   15285
      TabIndex        =   2
      Top             =   30
      Width           =   15285
      Begin RichTextLib.RichTextBox RichTextBox 
         Height          =   10425
         Left            =   150
         TabIndex        =   3
         Top             =   90
         Width           =   14985
         _ExtentX        =   26432
         _ExtentY        =   18389
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"frmChild.frx":08CA
      End
   End
End
Attribute VB_Name = "frmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim HeightDiff As Integer
Dim WidthDiff As Integer

Private Const FORMHEIGHT = 9000
Private Const FORMWIDTH = 10410

Sub Form_Load()
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
    
    Call refresh_richtext
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
Sub form_activate()
    Dim arraycount As Integer
    Dim i As Integer
    Dim doc As Integer
    
    arraycount = UBound(document)
    doc = current_form
    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    FState(doc).Dirty = True
End Sub



Private Sub form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Show the current form instance as deleted
    Dim doc As Integer
    
    doc = current_form
    FState(doc).deleted = True
    FState(doc).Dirty = False
    Call unload_document(doc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Show the current form instance as deleted
    Dim doc As Integer
    
    doc = current_form
    If doc <> -1 Then
        FState(doc).deleted = True
        FState(doc).Dirty = False
    End If
End Sub




















