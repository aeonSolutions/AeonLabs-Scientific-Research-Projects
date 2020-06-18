Attribute VB_Name = "Core"
' User-defined type to store information about child forms
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

Public last_window As String

Public FState()  As FormState           ' Array of user-defined types
Public document() As New frmChild
Public pf_graph() As New frmgraph
Public ri_graph() As New frmgraphbeta
Public fMainForm As frmMain
Public tipo As String


'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' Public functions used for disable de X button on the form
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
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


'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' support code for the execCmd sub
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' Support sub and vars for the sub copy_file
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

' API Declares for this module
Private Declare Function ReadFile Lib "kernel32" _
    (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
' Create a file handle on the local system
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" _
    (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
' Writes to a local file
Private Declare Function WriteFile Lib "kernel32" _
    (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Long) As Long

' Constants
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_ALWAYS = 2
Private Const CREATE_NEW = 1
Private Const OPEN_ALWAYS = 4
Private Const OPEN_EXISTING = 3
Private Const FILE_BEGIN = 0

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' Sub for copying files
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Public Sub Copy_file(stroriginalFile As String, strcopyfile As String)

Const bufferLen As Long = 1024 ' <<< Specify the amount of data that is transferred per iteration

Dim objFSO As Object, objFile As Object
Dim hLocalFile As Long, hNewFile As Long
Dim buffer As String, bytesRead As Long, bytesWritten As Long, bytesTransferred As Long
Dim boolCancel As Boolean

' Generate an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
' Then obtain the original file
Set objFile = objFSO.GetFile(stroriginalFile)

' Then set the Max property of the progress to the size of the file to be copied
frm_perform_calculations.pbar.Max = objFile.Size

' Then obtain the appropriate handle to the File to be copied
hLocalFile = CreateFile(stroriginalFile, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_ALWAYS, 0, 0)

If hLocalFile <> 0 Then
    ' Then obtain a handle to the new file
    hNewFile = CreateFile(strcopyfile, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, CREATE_ALWAYS, 0, 0)
    
    ' Set the buffer
    buffer = Space(bufferLen)
    
    If hNewFile <> 0 Then
        ' If OK - then proceed to open Read 'chunks' from the original file to the new file
        Do
            If ReadFile(hLocalFile, ByVal buffer, bufferLen, bytesRead, ByVal 0&) Then
                If WriteFile(hNewFile, ByVal buffer, bytesRead, bytesWritten, ByVal 0&) Then
                    ' Keep a tab on what's been downloaded
                    bytesTransferred = bytesTransferred + bytesWritten
                End If
            Else
                ' If there is no more to read, then exit
                boolCancel = True
            End If
        
            ' Update the transfer position by setting the progress bar's value to the current
            ' amount of bytes transferred
            frm_perform_calculations.pbar.Value = bytesTransferred
        
        Loop While bytesRead = bufferLen And Not boolCancel
    End If
End If


' Reset the Progress Bar value
frm_perform_calculations.pbar.Value = 0

' When completed - clean up objects / file handles
Set objFSO = Nothing
Set objFile = Nothing
CloseHandle (hLocalFile)
CloseHandle (hNewFile)
End Sub

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' Sub for making a msdos call - execute the core simulation program
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

   Public Function ExecCmd(ByRef cmdline)
      Dim proc As PROCESS_INFORMATION
      Dim Start As STARTUPINFO
      Dim ret&
      
      ' Initialize the STARTUPINFO structure:
      Start.cb = Len(Start)

      ' Start the shelled application:
      ret& = CreateProcessA(vbNullString, cmdline, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, vbNullString, Start, proc)

      ' Wait for the shelled application to finish:
         ret& = WaitForSingleObject(proc.hProcess, INFINITE)
         Call GetExitCodeProcess(proc.hProcess, ret&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = ret&
   End Function

'----------------------------------------------------------------------
'Description : Extracts the filename from a path
'Returns     : Returns the extracted filename, or the original string if no path exists
Public Function GetFile(ByVal PathAndFile As String) As String
 Dim r() As String
 If Len(PathAndFile) Then
  r() = Split(PathAndFile, "\")
  GetFile = r(UBound(r))
 End If
End Function '(Public) Function GetFile () As String

'----------------------------------------------------------------------
'Description : Removes the filename from path
'Returns     : Returns the path minus it's filename
Public Function GetPath(ByVal filename As String) As String
 Dim r() As String, p As String
 Dim i
 If InStr(filename, "\") Then
  r() = Split(filename, "\")
  For i = 0 To UBound(r) - 1
   p = p + r(i) + "\"
  Next
 End If
 GetPath = p
End Function '(Public) Function GetPath () As String

' finds a free index in the document array
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
    ReDim Preserve doc_props(arraycount + 1)
    ReDim Preserve pf_graph(arraycount + 1)
    ReDim Preserve ri_graph(arraycount + 1)
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
        ReDim pf_graph(1)
        ReDim ri_graph(1)
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
    FState(fIndex).Dirty = True
    FState(fIndex).saved = False
    FState(fIndex).newname = False
    FState(fIndex).calculated = False
    FState(fIndex).deleted = False
    FState(fIndex).values = False
    With doc_props(fIndex)
        .frm_ca_board3_values.values = False
        .frm_ca_board1_values.values = False
        .frm_ca_board2_values.values = False
        Call refresh_lista(fIndex)
    End With
    
    document(fIndex).Show
    FileNew = fIndex
End Function

' This function disables the X-button on a
' form, to keep the user from closing a form
'that way, but keeps the min & max buttons.
Public Sub DisableX(frm As Form)
  Dim hMenu As Long, nCount As Long
  hMenu = GetSystemMenu(frm.hwnd, 0)
  nCount = GetMenuItemCount(hMenu)
  Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
  DrawMenuBar frm.hwnd
End Sub

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


' Inserts a delay in seconds
Public Sub delay(tDelay As Double) 'in seconds
    Dim dTimer As Double
    
    dTimer = Timer
    While Timer < dTimer + tDelay
        DoEvents
    Wend
End Sub


Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Call delay(1)
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash
    fMainForm.Show
End Sub

