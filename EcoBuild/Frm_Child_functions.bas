Attribute VB_Name = "Frm_Child_functions"
Option Explicit
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' User-defined type to store information about child forms
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Type FormState
    deleted As Boolean
    Dirty As Boolean
    calculated As Boolean
    saved As Boolean
    newname As Boolean
    path As String
    name As String
    values As Boolean
    count As Integer
    db_pos As Integer
End Type

Public FState()  As FormState           ' Array of user-defined types
Public document() As New frmChild

Public frm_global() As New frm_graph_global
Public frm_co2() As New frm_graph_co2
Public frm_energy() As New frm_graph_energy
Public frm_so2() As New frm_graph_so2
Public frm_nox() As New frm_graph_nox
Public frm_structure() As New frm_graph_structure
Public frm_water() As New frm_graph_water
Public frm_report() As New frm_report_analysis

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' finds a free index in the document array and returns it's position in the array
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Function FindFreeIndex() As Integer
    Dim i As Integer
    Dim arraycount As Integer

    arraycount = UBound(document)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For i = 1 To arraycount
        If FState(i).deleted Then
            FindFreeIndex = i
            FState(i).deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.
    ReDim Preserve document(arraycount + 1)
    ReDim Preserve FState(arraycount + 1)
    ReDim Preserve frm_water(arraycount + 1)
    ReDim Preserve frm_structure(arraycount + 1)
    ReDim Preserve frm_co2(arraycount + 1)
    ReDim Preserve frm_so2(arraycount + 1)
    ReDim Preserve frm_nox(arraycount + 1)
    ReDim Preserve frm_global(arraycount + 1)
    ReDim Preserve frm_energy(arraycount + 1)
    ReDim Preserve frm_report(arraycount + 1)
    ReDim Preserve doc_props(arraycount + 1)

    FindFreeIndex = UBound(document)
End Function

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' creates a new frmchild form and returns it's position in the array
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Public Function FileNew() As Integer
    Dim fIndex As Integer
    Dim arraycount As Integer
    Dim i As Integer
    
    On Error Resume Next
    arraycount = UBound(document)

    If Err <> 0 Then
        ReDim document(1)
        
        ReDim frm_water(1)
        ReDim frm_structure(1)
        ReDim frm_co2(1)
        ReDim frm_so2(1)
        ReDim frm_nox(1)
        ReDim frm_global(1)
        ReDim frm_energy(1)
        ReDim frm_report(1)
        ReDim doc_props(1)

        ReDim FState(1)
        ReDim doc_props(1)
        document(1).Tag = "1"
        document(1).Caption = "Document nº" & Str(document(1).Tag)
        FState(1).Dirty = True
        FState(1).newname = False
        FState(1).saved = False
        FState(1).deleted = False
        FState(1).calculated = False
        FState(1).values = False
        document(1).Show
        FileNew = 1
        Exit Function
    End If
    ' Cycle through the document array
    For i = 0 To arraycount
         FState(i).Dirty = False
    Next
    ' Find the next available index and show the child form.
    fIndex = FindFreeIndex()
    document(fIndex).Tag = fIndex
    document(fIndex).Caption = "Document nº" & Str(fIndex)
    FState(fIndex).Dirty = True
    FState(fIndex).saved = False
    FState(fIndex).newname = False
    FState(fIndex).calculated = False
    FState(fIndex).deleted = False
    FState(fIndex).values = False
    document(fIndex).Show
    FileNew = fIndex
End Function

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' unloads the current formchild and free it's position in the array
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Public Sub unload_document()
Dim doc As Integer
Dim tmp As VbMsgBoxResult
Dim j As Integer
Dim name As String
Dim path As String

doc = current_form()
If FState(doc).values Then
  If Not FState(doc).saved Then
    tmp = MsgBox("Save the Document ?", vbYesNoCancel + vbCritical, "Temperus")
    If tmp = vbCancel Then
      Exit Sub
    End If
    If tmp = vbYes Then
           ' Set CancelError is True
           document(doc).dialogs.CancelError = True
           On Error Resume Next
           ' Set flags
           document(doc).dialogs.Flags = cdlOFNHideReadOnly
           ' Set filters
           document(doc).dialogs.Filter = dialogs_filter
           ' Specify default filter
           document(doc).dialogs.FilterIndex = 2
           ' set the working directory the application dir
           document(doc).dialogs.InitDir = App.path
           ' Display the save dialog box
           document(doc).dialogs.ShowSave
           If Err.Number <> 0 Then
             Exit Sub
           End If
           ' get the name file and the path
           name = GetFile(document(doc).dialogs.filename)
           path = GetPath(document(doc).dialogs.filename)
         Call save_file(name, path, doc)
    End If
  End If
End If
FState(doc).deleted = True
Unload document(doc)
FState(doc).deleted = True
FState(doc).Dirty = False

End Sub

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' returns the position on the array of the current formchild
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Public Function current_form() As Integer
  Dim i As Integer
  Dim arraycount As Integer
    ' Cycle through the document array
    i = 1
    On Error Resume Next
    arraycount = UBound(document)
    If Err <> 0 Then
        current_form = -1
        Exit Function
    End If
    For i = 1 To arraycount
         If FState(i).Dirty Then
           current_form = i
           Exit Function
         End If
    Next
  If i > arraycount Then
    i = arraycount
  End If
  current_form = i
End Function

