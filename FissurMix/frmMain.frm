VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "[CODED]"
   ClientHeight    =   9570
   ClientLeft      =   1755
   ClientTop       =   2970
   ClientWidth     =   10110
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Duracon"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   480
      Top             =   0
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   9300
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12197
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "02-09-2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "14:42"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog Dialogs 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnufilebar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_simulation 
      Caption         =   "&Simulation"
      Begin VB.Menu menu_load_data 
         Caption         =   "Load Experimental Data"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu_load_prob_stress 
         Caption         =   "Load Probability Stresses"
      End
      Begin VB.Menu menu_run 
         Caption         =   "&Run Analysis"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnu_results 
      Caption         =   "&Results"
      Begin VB.Menu mnu_load 
         Caption         =   "Load"
         Begin VB.Menu menu_exp_data 
            Caption         =   "&Polynomial Load Curve"
         End
         Begin VB.Menu menu_stress_curve 
            Caption         =   "&Composite Stress Curve"
         End
         Begin VB.Menu menu_emodulus_curve 
            Caption         =   "&Composite Elastic Modulus curve"
         End
         Begin VB.Menu menu_loading_stresses 
            Caption         =   "&Loading Stresses"
         End
         Begin VB.Menu menu_final_stresses 
            Caption         =   "&Final Stresses"
         End
      End
      Begin VB.Menu mnu_lmed 
         Caption         =   "Mean Crack Spacing"
         Begin VB.Menu mnu_lsf 
            Caption         =   "&Lmed versus Stress (Homogen.)"
         End
         Begin VB.Menu mnu_lsf_n 
            Caption         =   "L&med versus Stress"
         End
         Begin VB.Menu mnu_lstrain 
            Caption         =   "Stress(Lmed) versus &Strain"
         End
         Begin VB.Menu mnu_lmed_crk_density 
            Caption         =   "Lmed and Crk &density versus Strain"
         End
         Begin VB.Menu mnu_lmed_vs_slmed 
            Caption         =   "Lmed versus Stress(Lmed)"
         End
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
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

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub menu_emodulus_curve_Click()
Dim doc As Integer

doc = current_form
graph_type = "Composit Elastic Modulus Curve"
frm_exp_data(doc).Show 1

End Sub

Private Sub menu_exp_data_Click()
Dim doc As Integer

doc = current_form
graph_type = "Polynomial Load Curve"
frm_exp_data(doc).Show 1

End Sub


Private Sub menu_final_stresses_Click()
Dim doc As Integer

doc = current_form
graph_type = "Final Stresses"
frm_exp_data(doc).Show 1

End Sub

Private Sub menu_loading_stresses_Click()
Dim doc As Integer

doc = current_form
graph_type = "Loading Stresses"
frm_exp_data(doc).Show 1
End Sub

Private Sub menu_stress_curve_Click()
Dim doc As Integer

doc = current_form
graph_type = "Composit Stress Curve"
frm_exp_data(doc).Show 1
End Sub

Private Sub mnu_lmed_crk_density_Click()
Dim doc As Integer

doc = current_form
graph_type = "Lmed and Crk density versus Strain"
frm_exp_data(doc).Show 1
End Sub

Private Sub mnu_lmed_vs_slmed_Click()
Dim doc As Integer

doc = current_form
graph_type = "Lmed versus Stress(Lmed)"
frm_exp_data(doc).Show 1
End Sub

Private Sub mnu_lsf_Click()
Dim doc As Integer

doc = current_form
graph_type = "Lmed versus Stress (Homogen.)"
frm_exp_data(doc).Show 1
End Sub

Private Sub mnu_lsf_n_Click()
Dim doc As Integer

doc = current_form
graph_type = "Lmed versus Stress"
frm_exp_data(doc).Show 1

End Sub

Private Sub mnu_lstrain_Click()
Dim doc As Integer

doc = current_form
graph_type = "Stress(Lmed) versus Strain"
frm_exp_data(doc).Show 1

End Sub

Private Sub menu_run_Click()
Me.Enabled = False
Call Calcular
Me.Enabled = True
End Sub

Private Sub menu_load_data_Click()
Dim doc As Integer
Dim arraycount As Integer
Dim i As Integer
Dim tmp As Boolean

On Error Resume Next
arraycount = UBound(document)
If Err.Number <> 0 Then
    tmp = False
Else ' documents found - check if they really exist
    tmp = False
    For i = 1 To arraycount
        If Not FState(i).deleted Then
            tmp = True
            Exit For
        End If
    Next i
End If

If Not tmp Then
    MsgBox "You need to have at least one document open!", vbOKOnly + vbCritical, " Duracon "
    Exit Sub
End If

doc = current_form
frm_load_data.Show 1
End Sub
Private Sub mnu_load_prob_stress_Click()
    Dim tmp As VbMsgBoxResult
    Dim doc As Integer
    
    doc = current_form
    
    If doc_props(doc).elements_generated Then
        tmp = MsgBox("Probabilistic Stresses Already Generated. Generate New ones ?", vbYesNoCancel + vbCritical, " Temperus ")
        If tmp = vbCancel Or tmp = vbNo Then
            Exit Sub
        End If
    End If
    Call stress_load
End Sub
Private Sub MDIForm_Load()
    Me.Caption = App.Title
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    'FileNew
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub mnuEdit_Click()

End Sub

Private Sub mnuFileClose_Click()
Call unload_document(current_form)
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub


Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub


Private Sub mnuFileSaveAs_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  
  ' get the current form index
  cur_doc = current_form()
If cur_doc = False Then
    Exit Sub
End If
 If FState(cur_doc).deleted Then
       MsgBox "This message Should not appear! Look into IT PLEASE!", vbOKCancel, "Info"
     Exit Sub
  End If
    
    ' Set CancelError is True
    dialogs.CancelError = True
    On Error Resume Next
    ' Set flags
    dialogs.Flags = cdlOFNHideReadOnly
    ' Set filters
    dialogs.Filter = dialogs_filter
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
    name = GetFile(dialogs.filename)
    path = GetPath(dialogs.filename)
    Call save_file(name, path, cur_doc)
End Sub

Private Sub mnuFileSave_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  
  ' get the current form index
  cur_doc = current_form()
 If cur_doc = False Then
    Exit Sub
End If
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
    dialogs.Filter = dialogs_filter
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
    name = GetFile(dialogs.filename)
    path = GetPath(dialogs.filename)
    Call save_file(name, path, cur_doc)
    Exit Sub
  End If
  Call save_file(FState(cur_doc).name, FState(cur_doc).path, cur_doc)
End Sub

Private Sub mnuFileOpen_Click()
  Dim name As String
  Dim path As String
  Dim cur_doc As Integer
  Dim n As Integer
  Dim arraycount As Integer
  Dim i As Integer
  Dim old_doc As Integer
  Dim tmp As Integer
  Dim r() As String
  
  
' Set CancelError is True
  dialogs.CancelError = True
 On Error Resume Next
  ' Set flags
  dialogs.Flags = cdlOFNHideReadOnly And cdlOFNAllowMultiselect
  ' Set filters
  dialogs.Filter = dialogs_filter
  ' Specify default filter
  dialogs.FilterIndex = 2
  ' Display the open dialog box
  dialogs.ShowOpen
  If Err.Number = 32755 Then ' cancel was selected
    'MsgBox "error num:" & Str(Err.Number) & " Desc:" & Err.Description, vbOKCancel, "Info"
    Exit Sub
  End If
  ' get the name file and the path
name = GetFile(dialogs.filename)
  path = GetPath(dialogs.filename)
   
  r() = Split(name, ".")
  If r(UBound(r)) <> filename_extension Then
    MsgBox "Invalid File Type. You must select a Valid *." & filename_extension & " file.", vbOK, "Info"
    Exit Sub
  End If
   
  ' creates a new Child Form  and get the current form index
  On Error Resume Next
  arraycount = UBound(document)
  If Err <> 0 Or arraycount = -1 Then
        cur_doc = FileNew()
  Else
    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    old_doc = current_form
    If old_doc <> -1 Then  'there are child forms open
      If FState(old_doc).values = False Then
        FState(old_doc).deleted = True
        FState(old_doc).Dirty = False
      End If
    End If
    cur_doc = FileNew()
  End If
  ' change to the selected directory
  ChDir path
    Open name For Random As #1 Len = Len(doc_props(cur_doc))
    Get #1, 1, doc_props(cur_doc)
    Close #1
 
  document(cur_doc).Caption = name
  FState(cur_doc).path = path
  FState(cur_doc).name = name
  FState(cur_doc).saved = True
  FState(cur_doc).newname = True
  FState(cur_doc).calculated = False
  FState(cur_doc).values = True
  
  Call refresh_richtext
  
End Sub

Private Sub mnuFileNew_Click()
    FileNew
End Sub

Private Sub Timer_Timer()
Dim doc As Integer
Dim tmp As Boolean
Dim arraycount As Integer
Dim i As Integer

Dim path As String
On Error Resume Next
arraycount = UBound(document)
If Err.Number <> 0 Then
    tmp = False
Else ' documents found - check if they really exist
    tmp = False
    For i = 1 To arraycount
        If Not FState(i).deleted Then
            tmp = True
            Exit For
        End If
    Next i
End If

If tmp Then ' documents found
    mnu_simulation.Enabled = True
    mnuFileClose.Enabled = True
    doc = current_form
    If FState(doc).values Then
        mnu_results.Enabled = True
    End If
    If FState(doc).calculated Then
        menu_final_stresses.Enabled = True
        menu_loading_stresses.Enabled = True
        mnu_lsf.Enabled = True
        mnu_lsf_n.Enabled = True
        mnu_lstrain.Enabled = True
        mnu_lmed_crk_density.Enabled = True
        mnu_lmed_vs_slmed.Enabled = True
    End If
Else
    mnu_lmed_vs_slmed.Enabled = False
    mnu_lmed_crk_density.Enabled = False
    mnu_lstrain.Enabled = False
    mnu_lsf_n.Enabled = False
    mnu_lsf.Enabled = False
    mnu_simulation.Enabled = False
    mnuFileClose.Enabled = False
    mnu_results.Enabled = False
    menu_final_stresses.Enabled = False
    menu_loading_stresses.Enabled = False
End If


End Sub

