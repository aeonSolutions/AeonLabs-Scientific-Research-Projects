Attribute VB_Name = "General_mod"

' User-defined type to store information about child forms
Type FormState
    deleted As Boolean
    Dirty As Boolean
    calculated As Boolean
    Conta As Integer
    saved As Boolean
    newname As Boolean
    path As String
    name As String
End Type

Public FState()  As FormState           ' Array of user-defined types
Public document() As New frmDocument
Public fMainForm As frmMain
Public tipo As String

' Public functions used for disable de X button on the form
Public Const MF_BYPOSITION = &H400&
Public Const MF_DISABLED = &H2&
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
'end of the public fucntions for disable X

'----------------------------------------------------------------------
'Description : Extracts the filename from a path
'----------------------------------------------------------------------
'Returns     : Returns the extracted filename, or the original string if no path exists
'----------------------------------------------------------------------
Public Function GetFile(ByVal PathAndFile As String) As String
 Dim R() As String
 If Len(PathAndFile) Then
  R() = Split(PathAndFile, "\")
  GetFile = R(UBound(R))
 End If
End Function '(Public) Function GetFile () As String

'----------------------------------------------------------------------
'Description : Removes the filename from path
'----------------------------------------------------------------------
'Returns     : Returns the path minus it's filename
'----------------------------------------------------------------------
Public Function GetPath(ByVal Filename As String) As String
 Dim R() As String, p As String
 Dim i
 If InStr(Filename, "\") Then
  R() = Split(Filename, "\")
  For i = 0 To UBound(R) - 1
   p = p + R(i) + "\"
  Next
 End If
 GetPath = p
End Function '(Public) Function GetPath () As String

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
    FindFreeIndex = UBound(document)
End Function

Public Function FileNew() As Integer
    Dim fIndex As Integer
    Dim arraycount As Integer
    Dim i As Integer
    
    On Error Resume Next
    arraycount = UBound(document)

    If Err <> 0 Then
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
        document(1).Show
        FileNew = 1
        Exit Function
    End If
    ' Cycle through the document array
    For i = 1 To arraycount
         FState(i).Dirty = False
    Next
    ' Find the next available index and show the child form.
    fIndex = FindFreeIndex()
    document(fIndex).Tag = fIndex
    document(fIndex).Caption = "Document nº" & Str(fIndex)
    FState(fIndex).Conta = 1
    FState(fIndex).Dirty = True
    FState(fIndex).saved = False
    FState(fIndex).newname = False
    FState(fIndex).calculated = False
    FState(fIndex).deleted = False
    document(fIndex).Show
    FileNew = fIndex
End Function


Public Sub delay(tDelay As Double) 'in seconds
    Dim dTimer As Double
    
    dTimer = Timer
    While Timer < dTimer + tDelay
        DoEvents
    Wend
End Sub

'* Purpose    : Disables the close button ('X') on form.    *
'* Description: This function disables the X-button on a    *
'*            : form, to keep the user from closing a form  *
'*            : that way, but keeps the min & max buttons.  *
Public Sub DisableX(frm As Form)
  Dim hMenu As Long, nCount As Long
  'Get handle to system menu
  hMenu = GetSystemMenu(frm.hwnd, 0)
  'Get number of items in menu
  nCount = GetMenuItemCount(hMenu)
  'Remove last item from system menu (last item is 'Close')
  Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
  'Redraw menu
  DrawMenuBar frm.hwnd
End Sub
Public Sub savefile(ByRef name, ByRef path, cur_doc As Integer)
  Dim n As Integer
  ' change to the selected directory
  ChDir path

  ReDim material(FState(cur_doc).Conta - 1)
  n = FState(cur_doc).Conta - 1
  For i = 1 To n
     material(i).num_mats = n
     document(cur_doc).lista.row = i
     document(cur_doc).lista.col = 1
     material(i).l = CDbl(document(cur_doc).lista.Text) * units.l.conversion(units.l.selected)
     document(cur_doc).lista.col = 2
     material(i).area = CDbl(document(cur_doc).lista.Text) * units.area.conversion(units.area.selected)
     document(cur_doc).lista.col = 3
     material(i).k = CDbl(document(cur_doc).lista.Text) * units.k.conversion(units.k.selected)
     document(cur_doc).lista.col = 4
     material(i).b = CDbl(document(cur_doc).lista.Text) * units.b.conversion(units.b.selected)
     document(cur_doc).lista.col = 5
     material(i).te = CDbl(document(cur_doc).lista.Text) * units.te.conversion(units.te.selected)
     document(cur_doc).lista.col = 6
     material(i).td = CDbl(document(cur_doc).lista.Text) * units.td.conversion(units.td.selected)
     document(cur_doc).lista.col = 7
     material(i).n = CDbl(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 8
     material(i).e = CDbl(document(cur_doc).lista.Text) * units.e.conversion(units.e.selected)
     document(cur_doc).lista.col = 9
     material(i).alfa = CDbl(document(cur_doc).lista.Text)
     document(cur_doc).lista.col = 10
     material(i).q0 = CDbl(document(cur_doc).lista.Text) / units.q0.conversion(units.q0.selected)
  Next i
  
  Open name For Random As #1 Len = Len(material(1))
  With document(cur_doc)
    For i = 1 To FState(cur_doc).Conta - 1
      Put #1, i, material(i)
    Next i
  End With
 Close #1
 FState(cur_doc).saved = True
 document(cur_doc).Caption = name

End Sub

Public Function current_form() As Integer
  Dim i As Integer
  Dim arraycount As Integer
    ' Cycle through the document array
    i = 1
    arraycount = UBound(document)
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



Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Call delay(1)
    
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash
End Sub

