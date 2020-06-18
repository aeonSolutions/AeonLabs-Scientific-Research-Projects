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
    
End Type

Public FState()  As FormState           ' Array of user-defined types
Public document() As New frmChild


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
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.
    ReDim Preserve document(arraycount + 1)
    ReDim Preserve FState(arraycount + 1)
    
    ReDim Preserve frm_exp_data(arraycount + 1)
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
    If Err <> 0 Or arraycount = -1 Then
        
        ReDim frm_exp_data(1)
        ReDim frm_segment_cracks(1)
    
        ReDim document(1)
        ReDim FState(1)
        ReDim doc_props(1)

        document(1).Tag = "1"
        document(1).Caption = "Document nº1"
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
Public Sub unload_document(doc As Integer)
Dim tmp As VbMsgBoxResult
Dim j As Integer
Dim name As String
Dim path As String

If FState(doc).values Then
  If Not FState(doc).saved Then
    tmp = MsgBox("Save the Document ?", vbYesNoCancel + vbCritical, "Temperus")
    If tmp = vbCancel Then
      Exit Sub
    End If
    If tmp = vbYes Then
           ' Set CancelError is True
           document(doc).Dialogs.CancelError = True
           On Error Resume Next
           ' Set flags
           document(doc).Dialogs.Flags = cdlOFNHideReadOnly
           ' Set filters
           document(doc).Dialogs.Filter = dialogs_filter
           ' Specify default filter
           document(doc).Dialogs.FilterIndex = 2
           ' set the working directory the application dir
           document(doc).Dialogs.InitDir = App.path
           ' Display the save dialog box
           document(doc).Dialogs.ShowSave
           If Err.Number <> 0 Then
             Exit Sub
           End If
           ' get the name file and the path
           name = GetFile(document(doc).Dialogs.filename)
           path = GetPath(document(doc).Dialogs.filename)
         Call save_file(name, path, doc)
    End If
  End If
End If
FState(doc).deleted = True
Unload document(doc)
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
    current_form = -5
    For i = 1 To arraycount
         If FState(i).deleted = False Then
           current_form = -3
         End If
    Next
    If current_form = -5 Then
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

