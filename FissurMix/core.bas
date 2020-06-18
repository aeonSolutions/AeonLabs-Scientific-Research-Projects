Attribute VB_Name = "Core"
Option Explicit

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


'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' This function disables the X-button on a form, to keep the user
' from closing a form that way, but keeps the min & max buttons.
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Public Sub DisableX(frm As Form)
  Dim hMenu As Long, nCount As Long
  hMenu = GetSystemMenu(frm.hwnd, 0)
  nCount = GetMenuItemCount(hMenu)
  Call RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)
  DrawMenuBar frm.hwnd
End Sub



'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' Inserts a delay in seconds
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
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

