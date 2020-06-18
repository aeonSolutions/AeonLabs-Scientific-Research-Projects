Attribute VB_Name = "Locale_vars_conversion"
Option Explicit


'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' this func converts correctly a string value into a numeric value
' according to the LOCALE decimal separator
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Public Function convert_type(s As String) As Double
If decimal_separator = "," Then
    s = Replace(s, ",", ".")
    convert_type = Val(s)
Else
    convert_type = Val(s)
End If
End Function

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' this func converts correctly a string value into another string value
' according to the LOCALE decimal separator
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Public Function str2str(s As String) As String
If decimal_separator = "," Then
    str2str = Replace(s, ".", ",")
Else
    str2str = s
End If

End Function

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' this func returns the decimal separator
' according to the LOCALE setting defined in the computer
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Public Function decimal_separator() As String
  ' instanciate the object and return some of the system format settings....
  Dim oGetFormats As cGetLocalFormats
  Set oGetFormats = New cGetLocalFormats
  
  With oGetFormats
    decimal_separator = .NumericDecimalSeparator
  End With
  
  Set oGetFormats = Nothing


End Function

