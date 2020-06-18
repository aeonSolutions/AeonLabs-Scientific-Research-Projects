VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "FissurMax"
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
   Begin MSComctlLib.StatusBar StatusBar 
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
            TextSave        =   "12-07-2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "20:38"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
   Begin VB.Menu menu_material 
      Caption         =   "Material"
      Begin VB.Menu menu_material_generate 
         Caption         =   "&Load Probability Stresses"
         Shortcut        =   ^L
      End
      Begin VB.Menu menu_MaterialRun 
         Caption         =   "&Run Statistic Analisys"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu menu_View 
      Caption         =   "&View"
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
    
    ReDim document(1)
    ReDim FState(1)
    document(1).Tag = 1
    document(1).Caption = "Document nº" & Str(document(1).Tag)
    FState(1).Dirty = True
    FState(1).newname = False
    FState(1).saved = False
    FState(1).Conta = 1
    FState(1).deleted = False
    FState(1).calculated = False
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
            document(i).chart.EditCopy
            Me.StatusBar.Panels(1).Text = "Frequency Graph copied..."

        End If
        If .Tab = 2 Then
            document(i).chart3.EditCopy
            Me.StatusBar.Panels(1).Text = "Cumulated Frequency Graph copied..."

        End If
        If .Tab = 3 Then
            document(i).chart2.EditCopy
            Me.StatusBar.Panels(1).Text = "Final Stresses Graph copied..."
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
If Not FState(i).saved Then
  If Not FState(i).saved Then
    tmp = MsgBox("Save the Document ?", vbYesNoCancel + vbCritical, "Temperus")
    If tmp = vbCancel Then
      Exit Sub
    End If
    If tmp = vbYes Then
           ' Set CancelError is True
           Dialogs.CancelError = True
           On Error Resume Next
           ' Set flags
           Dialogs.Flags = cdlOFNHideReadOnly
           ' Set filters
           Dialogs.Filter = "All Files (*.*)|*.*|FissurMax Files" & _
           "(*.fmx)|*.fmx"
           ' Specify default filter
           Dialogs.FilterIndex = 2
           Dialogs.InitDir = App.path
           ' Display the save dialog box
           Dialogs.ShowSave
           If Err.Number <> 0 Then
             Exit Sub
           End If
           ' get the name file and the path
           name = GetFile(Dialogs.Filename)
           path = GetPath(Dialogs.Filename)
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
  Dialogs.CancelError = True
  On Error Resume Next
  ' Set flags
  Dialogs.Flags = cdlOFNHideReadOnly
  ' Set filters
  Dialogs.Filter = "All Files (*.*)|*.*|Text Files" & _
  "(*.csv)|*.csv"
  ' Specify default filter
  Dialogs.FilterIndex = 2
  ' Display the save dialog box
  Dialogs.ShowSave
  If Err.Number <> 0 Then
    Exit Sub
  End If
  ' get the name file and the path
  name = GetFile(Dialogs.Filename)
  path = GetPath(Dialogs.Filename)
  path = GetPath(Dialogs.Filename)

  ' change to the selected directory
  ChDir path
    num_elements = CDbl(document(cur_doc).divisoes_txt.Text)

  ReDim arrays(num_elements + 10)
  'arrays(1) = "nº;L (mm);Area (mm2);K (W/mºC);q0 (W/m3);T1 (ºC);G0 (W/mm2);n;E (GPa);alfa"
  'arrays(2)=
  j = 1
  Load percent
  percent.Show
  percent.SetFocus
  Call DisableX(percent)
  percent.overall_txt.Caption = "Exporting results table... One moment, please!"
  percent.overall_pbar.Max = num_elements + 11
  Call delay(0.02)

  arrays(j) = "RESULTADOS;"
  j = j + 1
  For i = 0 To document(cur_doc).results.Rows
   percent.overall_pbar.Value = i
   With document(cur_doc)
        .results.Row = i
        .results.Col = 0
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 1
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 2
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 3
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 4
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 5
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 6
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 7
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        .results.Col = 8
        arrays(j + i) = arrays(j + i) & .results.Text & ";"
        
    End With
Next i
Open Dialogs.Filename For Output As #1
For i = 1 To UBound(arrays)
    Print #1, arrays(i)
Next i
Close #1
percent.Hide
Unload percent

Me.StatusBar.Panels(1).Text = "File Successfully exported..."
End Sub

Private Sub menu_FileNew_Click()
    FileNew
    Me.StatusBar.Panels(1).Text = "New Document created..."
    
End Sub

Private Sub menu_FileOpen_Click()
      Dim name As String
      Dim path As String
      Dim cur_doc As Integer
      Dim n As Integer
      Dim arraycount As Integer
      Dim i As Integer
      
    
    ' Set CancelError is True
      Dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    Dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialogs.Filter = "All Files (*.*)|*.*|FissurMax Files" & _
    "(*.fmx)|*.fmx"
    ' Specify default filter
    Dialogs.FilterIndex = 2
    ' Display the save dialog box
    Dialogs.ShowOpen
    If Err.Number <> 0 Then
      Exit Sub
    End If
    ' get the name file and the path
    name = GetFile(Dialogs.Filename)
    path = GetPath(Dialogs.Filename)
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
    Open name For Random As #1 Len = Len(material)
    FState(cur_doc).Conta = 1
    Get #1, 1, material
    Close #1
    With material
        document(cur_doc).tf_txt = CStr(.tf)
        document(cur_doc).l_txt = CStr(.l)
        document(cur_doc).ef_txt = CStr(.ef)
        document(cur_doc).vf_txt = CStr(.vf)
        document(cur_doc).ts_txt = CStr(.ts)
        document(cur_doc).vs_txt = CStr(.vs)
        document(cur_doc).es_txt = CStr(.es)
        document(cur_doc).sigma_txt = CStr(.sigma)
        document(cur_doc).delta_txt = CStr(.delta)
        document(cur_doc).rs_txt = CStr(.rs)
        document(cur_doc).m_txt = CStr(.m)
        document(cur_doc).sigma_weib_txt = CStr(.sigma_weib)
        document(cur_doc).n_txt = CStr(.n)
        document(cur_doc).divisoes_txt = CStr(.divisoes)
        document(cur_doc).segments_txt.Text = CStr(.segments)
    End With
  
  arraycount = UBound(document)
  ' Cycle through the document array
  For i = 1 To arraycount
         FState(i).Dirty = False
  Next i
  FState(cur_doc).Dirty = True
  Me.StatusBar.Panels(1).Text = "File Successfully open..."

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
    Dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    Dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialogs.Filter = "All Files (*.*)|*.*|FissurMax Files" & _
    "(*.fmx)|*.fmx"
    ' Specify default filter
    Dialogs.FilterIndex = 2
    ' set the working directory the application dir
    Dialogs.InitDir = App.path
    ' Display the save dialog box
    Dialogs.ShowSave
    If Err.Number <> 0 Then
      Exit Sub
    End If
    ' get the name file and the path
    name = GetFile(Dialogs.Filename)
    path = GetPath(Dialogs.Filename)
    Call savefile(name, path, cur_doc)
    Exit Sub
  End If
  Call savefile(FState(cur_doc).name, FState(cur_doc).path, cur_doc)
  Me.StatusBar.Panels(1).Text = "File Successfully saved..."
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
    Dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    Dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    Dialogs.Filter = "All Files (*.*)|*.*|FissurMax Files" & _
    "(*.fmx)|*.fmx"
    ' Specify default filter
    Dialogs.FilterIndex = 2
    ' set the working directory the application dir
    Dialogs.InitDir = App.path
    ' Display the save dialog box
    Dialogs.ShowSave
    If Err.Number <> 0 Then
      Exit Sub
    End If
    ' get the name file and the path
    name = GetFile(Dialogs.Filename)
    path = GetPath(Dialogs.Filename)
    Call savefile(name, path, cur_doc)
    Me.StatusBar.Panels(1).Text = "File Successfully saved..."
End Sub

Private Sub menu_HelpAbout_Click()
Load frmAbout
frmAbout.Show 1
End Sub



Private Sub menu_HelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Sorry! Under Development.", vbInformation, Me.Caption
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
        MsgBox "Sorry! Under Development.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub menu_material_generate_Click()
    Dim tmp As VbMsgBoxResult

    
    If generated Then
        tmp = MsgBox("Probabilistic Stresses Already Generated. Generate New ones ?", vbYesNoCancel + vbCritical, " Temperus ")
        If tmp = vbCancel Or tmp = vbNo Then
            Exit Sub
        End If
    End If
    
    Call stress_load
        
End Sub

Private Sub menu_MaterialRun_Click()
    Me.StatusBar.Panels(1).Text = "Running Statistic Analisys...One moment please"
    Call Calcular
    Me.StatusBar.Panels(1).Text = "Statistic Analisys Complete..."
End Sub



Private Sub menu_ViewStatusBar_Click()
    menu_ViewStatusBar.Checked = Not menu_ViewStatusBar.Checked
    StatusBar.Visible = menu_ViewStatusBar.Checked

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
