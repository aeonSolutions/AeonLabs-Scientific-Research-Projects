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
Public Declare Function GetSystemMenu Lib "user32" ( _
                                   ByVal hwnd As Long, _
                                   ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" ( _
                                   ByVal hMenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" ( _
                                   ByVal hMenu As Long, _
                                   ByVal nPosition As Long, _
                                   ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" ( _
                                   ByVal hwnd As Long) As Long
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

Public Function current_form() As Integer
  Dim i As Integer
  Dim arraycount As Integer
    ' Cycle through the document array
    i = 1
    arraycount = UBound(document)
    For i = 1 To arraycount
         If FState(i).Dirty Then
           Exit For
         End If
    Next
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

