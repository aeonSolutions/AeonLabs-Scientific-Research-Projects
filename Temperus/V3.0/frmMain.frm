VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Temperus"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dialogs 
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2593
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "02-06-2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "11:30"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menu_File 
      Caption         =   "&File"
      Begin VB.Menu menu_FileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu menu_FileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu menu_FileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu menu_FileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu menu_FileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu menu_FileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu menu_FileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_FileExport 
         Caption         =   "&Export"
      End
      Begin VB.Menu menu_FileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_FilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu menu_FileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_FileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menu_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu menu_EditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu menu_Material 
      Caption         =   "Material"
      Begin VB.Menu menu_MaterialInsert 
         Caption         =   "&Manage"
         Shortcut        =   ^M
      End
      Begin VB.Menu menu_MaterialRun 
         Caption         =   "&Calculate"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu menu_View 
      Caption         =   "&View"
      Begin VB.Menu menu_ViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu menu_ViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu menu_Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu menu_WindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu menu_WindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu menu_WindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
   End
   Begin VB.Menu menu_Help 
      Caption         =   "&Help"
      Begin VB.Menu menu_HelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu menu_HelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu menu_HelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu menu_HelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Dim num_elements As Integer
Dim i As Integer



Private Sub MDIForm_Load()
    Show
    ' Always set the working directory to the directory containing the application.
    ChDir App.path
    ' Initialize the document form array, and show the first document.
    
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub menu_EditCopy_Click()
Dim i As Integer

i = current_form
With document(i)
    With .SSTab
        If .Tab = 0 Then
            
        End If
        If .Tab = 1 Then
            document(i).temperature_big_chart.EditCopy
        End If
        If .Tab = 2 Then
            document(i).displacement_big_chart.EditCopy
        End If
        If .Tab = 3 Then
            document(i).tension_big_chart.EditCopy
        End If
        
    End With
End With
End Sub

Private Sub menu_FileClose_Click()
Dim i As Integer
Dim tmp As VbMsgBoxResult
Dim name As String
Dim path As String


i = current_form()
If FState(i).Conta > 1 Then
  If Not FState(i).saved Then
    tmp = MsgBox("Save the Document ?", vbYesNoCancel + vbCritical, "Temperus")
    If tmp = vbCancel Then
      Exit Sub
    End If
    If tmp = vbYes Then
           ' Set CancelError is True
           dialogs.CancelError = True
           On Error Resume Next
           ' Set flags
           dialogs.Flags = cdlOFNHideReadOnly
           ' Set filters
           dialogs.Filter = "All Files (*.*)|*.*|Temperus Files" & _
           "(*.tps)|*.tps"
           ' Specify default filter
           dialogs.FilterIndex = 2
           ' set the working directory the application dir
           dialogs.InitDir = App.path
           ' Display the save dialog box
           dialogs.ShowSave
           If Err.Number <> 0 Then
             Exit Sub
           End If
           ' get the name file and the path
           name = GetFile(dialogs.Filename)
           path = GetPath(dialogs.Filename)
         Call savefile(name, path, i)
    End If
  End If
End If
FState(i).deleted = True
Unload document(i)
End Sub

Private Sub menu_FileExit_Click()
   Unload Me
End Sub

Private Sub menu_FileExport_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim arrays() As String
  Dim i As Integer
  Dim j As Integer


  
  ' get the current form index
  cur_doc = current_form()
  If FState(cur_doc).deleted Then
     Exit Sub
  End If
  If FState(cur_doc).calculated = False Then
        MsgBox "Calculate First!", vbCritical, " Temperus "
    Exit Sub
  End If
  
  ' Set CancelError is True
  dialogs.CancelError = True
  On Error Resume Next
  ' Set flags
  dialogs.Flags = cdlOFNHideReadOnly
  ' Set filters
  dialogs.Filter = "All Files (*.*)|*.*|Text Files" & _
  "(*.csv)|*.csv"
  ' Specify default filter
  dialogs.FilterIndex = 2
  ' Display the save dialog box
  dialogs.ShowSave
  If Err.Number <> 0 Then
    Exit Sub
  End If
  ' get the name file and the path
  name = GetFile(dialogs.Filename)
  path = GetPath(dialogs.Filename)

  ' change to the selected directory
  ChDir path

  ReDim material(FState(cur_doc).Conta - 1)
  ReDim arrays(document(cur_doc).results.Rows + 8 + FState(cur_doc).Conta)
  arrays(1) = "nº;L (mm);Area (mm2);K (W/mºC);q0 (W/m3);T1 (ºC);G0 (W/mm2);n;E (GPa);alfa"
  For i = 0 To FState(cur_doc).Conta - 1
    With document(cur_doc)
        .lista.row = i
        .lista.col = 0
        arrays(i + 1) = Str(i) & ";"
        .lista.col = 1
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 2
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 3
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 4
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 5
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 6
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 7
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 8
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 9
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
        .lista.col = 10
        arrays(i + 1) = arrays(i + 1) & .lista.Text & ";"
    End With
  Next i
    j = FState(cur_doc).Conta + 2
    arrays(j) = "RESULTADOS;"
  j = FState(cur_doc).Conta + 4
  For i = 0 To document(cur_doc).results.Rows
    With document(cur_doc)
        .results.row = i
        .results.col = 0
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.col = 1
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.col = 2
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.col = 3
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.col = 4
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        
    End With
Next i
Open dialogs.Filename For Output As #1
For i = 1 To UBound(arrays)
    Print #1, arrays(i)
Next i
Close #1
End Sub

Private Sub menu_FileNew_Click()
    FileNew
End Sub

Private Sub menu_FileOpen_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  Dim arraycount As Integer
  Dim i As Integer
  

' Set CancelError is True
  dialogs.CancelError = True
  On Error Resume Next
  ' Set flags
  dialogs.Flags = cdlOFNHideReadOnly
  ' Set filters
  dialogs.Filter = "All Files (*.*)|*.*|Temperus Files" & _
  "(*.tps)|*.tps"
  ' Specify default filter
  dialogs.FilterIndex = 2
  ' Display the save dialog box
  dialogs.ShowOpen
  If Err.Number <> 0 Then
    Exit Sub
  End If
  ' get the name file and the path
  name = GetFile(dialogs.Filename)
  path = GetPath(dialogs.Filename)
  ' creates a new Clid Form  and get the current form index
  arraycount = UBound(document)
  ' Cycle through the document array
  For i = 1 To arraycount
       FState(i).Dirty = False
  Next
  cur_doc = FileNew()
  ' change to the selected directory
  ChDir path
  
  document(cur_doc).Caption = name
  FState(cur_doc).path = path
  FState(cur_doc).name = name
  FState(cur_doc).saved = True
  FState(cur_doc).newname = True
  FState(cur_doc).calculated = False
  ReDim material(1)
  Open name For Random As #1 Len = Len(material(1))
  'n = LOF(1) / Len(material(1))
  Get #1, 1, material(1)
  n = material(1).num_mats
  FState(cur_doc).Conta = n + 1
  ReDim material(n)
  For i = 1 To n
    Get #1, i, material(i)
  Next i
  Close #1
  For i = 1 To n
     With document(cur_doc)
       .lista.row = i
       .lista.col = 0
       .lista.CellAlignment = 4
       .lista.Text = Str(i)
       .lista.col = 1
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).l / units.l.conversion(units.l.selected)))
       .lista.col = 2
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).area / units.area.conversion(units.area.selected)))
       .lista.col = 3
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).k / units.k.conversion(units.k.selected)))
       .lista.col = 4
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).b / units.b.conversion(units.b.selected)))
       .lista.col = 5
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).te / units.te.conversion(units.te.selected)))
       .lista.col = 6
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).td / units.td.conversion(units.td.selected)))
       .lista.col = 7
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).n))
       .lista.col = 8
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).e / units.e.conversion(units.e.selected)))
       .lista.col = 9
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).alfa))
       .lista.col = 10
       .lista.CellAlignment = 4
       .lista.Text = (CStr(material(i).q0 * units.q0.conversion(units.q0.selected)))

    End With
  Next i
  document(cur_doc).lista.Refresh
End Sub

Private Sub menu_FileSave_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  
  ' get the current form index
  cur_doc = current_form()
 If FState(cur_doc).deleted Then
       MsgBox "ERROR -This message Should not appear!", vbOKCancel, "Info"
     Exit Sub
  End If
  If FState(cur_doc).saved Then
    MsgBox "Document already Saved!", vbOKCancel, "Info"
    Exit Sub
  End If
  If Not FState(cur_doc).newname Then
    ' Set CancelError is True
    dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    dialogs.Filter = "All Files (*.*)|*.*|Temperus Files" & _
    "(*.tps)|*.tps"
    ' Specify default filter
    dialogs.FilterIndex = 2
    ' set the working directory the application dir
    dialogs.InitDir = App.path
    ' Display the save dialog box
    dialogs.ShowSave
    If Err.Number <> 0 Then
      Exit Sub
    End If
    ' get the name file and the path
    name = GetFile(dialogs.Filename)
    path = GetPath(dialogs.Filename)
    Call savefile(name, path, cur_doc)
    Exit Sub
  End If
  Call savefile(FState(cur_doc).name, FState(cur_doc).path, cur_doc)

End Sub

Sub menu_FileSaveAs_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  
  ' get the current form index
  cur_doc = current_form()
 If FState(cur_doc).deleted Then
       MsgBox "ERROR -This message Should not appear!", vbOKCancel, "Info"
     Exit Sub
  End If

    ' Set CancelError is True
    dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    dialogs.Filter = "All Files (*.*)|*.*|Temperus Files" & _
    "(*.tps)|*.tps"
    ' Specify default filter
    dialogs.FilterIndex = 2
    ' set the working directory the application dir
    dialogs.InitDir = App.path
    ' Display the save dialog box
    dialogs.ShowSave
    If Err.Number <> 0 Then
      Exit Sub
    End If
    ' get the name file and the path
    name = GetFile(dialogs.Filename)
    path = GetPath(dialogs.Filename)
    Call savefile(name, path, cur_doc)

End Sub

Private Sub menu_HelpAbout_Click()
Load frmAbout
frmAbout.Show 1
End Sub



Private Sub menu_MaterialInsert_Click()
Dim arraycount As Integer

    On Error Resume Next
    arraycount = UBound(document)
    If Err.Number <> 0 Then
       MsgBox "You need to have at least one document open!", vbOK + vbCritical, " Temperus "
       Exit Sub
    End If
    tipo = "new"
    Load frm_add_new_material
    frm_add_new_material.Show 1
End Sub

Private Sub menu_HelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub menu_HelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub
Private Sub menu_FilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dialogs
        .DialogTitle = "Print"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub


Private Sub menu_FilePageSetup_Click()
    On Error Resume Next
    With dialogs
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub
Private Sub menu_MaterialRun_Click()
    Dim i As Integer
    
    Me.Enabled = False
    Call calculus
    Me.Enabled = True
    i = current_form
    document(i).lista.SetFocus
    
End Sub


Private Sub menu_ViewOptions_Click()

End Sub

Private Sub menu_ViewStatusBar_Click()
    menu_ViewStatusBar.Checked = Not menu_ViewStatusBar.Checked
    sbStatusBar.Visible = menu_ViewStatusBar.Checked

End Sub

Private Sub menu_ViewToolbar_Click()
    menu_ViewToolbar.Checked = Not menu_ViewToolbar.Checked
   ' tbToolBar.Visible = menu_ViewToolbar.Checked

End Sub

Private Sub menu_WindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub menu_WindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub menu_WindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
