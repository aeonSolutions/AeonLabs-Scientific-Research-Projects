Attribute VB_Name = "modCenterform"
'**************************************************
' MODULE NAME : Centerform
' PURPOSE     : Centers a specified form.
' AUTHOR      : Ojie Maverick
' PARAMETERS  : FormName
' RETURNS     : none
'**************************************************
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Const conHwndTopaost = -1
Public Const SWP_NOSIZE = &H1

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Function Centerform(Frm As Form) As Boolean
Attribute Centerform.VB_Description = "Centers a specified form in the screen or within the mdi form."
On Error Resume Next
  Dim x As Integer
  Dim Y As Integer
  Dim rct As RECT

  'On Error GoTo errmsg
  If Frm.MDIChild Then
    GetClientRect GetParent(Frm.hWnd), rct
    x = ((rct.Right - rct.Left) * Screen.TwipsPerPixelY - _
        Frm.Width) * 0.5
    Y = ((rct.Bottom - rct.Top) * Screen.TwipsPerPixelX - _
            Frm.Height) * 0.5
  End If

  If Not Frm.MDIChild Then
    x = (Screen.Width - Frm.Width) * 0.5
    Y = (Screen.Height - Frm.Height) * 0.5
  End If

  Frm.Move x, Y
  Centerform = True

End Function

