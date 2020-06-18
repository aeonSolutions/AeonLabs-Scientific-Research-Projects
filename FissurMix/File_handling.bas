Attribute VB_Name = "File_handling"
Option Explicit

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
   
   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

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
' Sub for saving the input data in the current active frmchild
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Public Sub save_file(ByRef name, ByRef path, cur_doc As Integer)
  Dim doc As Integer
  Dim filename As String
  
  doc = current_form
  ChDir path
' Put the code for saving the file in the next line

ChDir path
filename = name
Open name For Random As #1 Len = Len(doc_props(doc))
    Put #1, 1, doc_props(doc)
Close #1

FState(doc).saved = True
document(doc).Caption = name
End Sub


