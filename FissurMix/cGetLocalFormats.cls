VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGetLocalFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Option Explicit



  ' enum of the format constants we will use
  Private Enum LOCAL_FORMATS
    eListSeparator = LOCALE_SLIST
    eNumericDecimalSeparator = LOCALE_SDECIMAL
    eCurrencySymbol = LOCALE_SCURRENCY
    eCurrencyDecimalSeparator = LOCALE_SMONDECIMALSEP
    eCurrencyThousandsSeparator = LOCALE_SMONTHOUSANDSEP
    eCurrencyPrefixOrSuffixMode = LOCALE_ICURRENCY
    eDateSeparator = LOCALE_SDATE
    eShortDateFormat = LOCALE_IDATE
    eDayLeadingZero = LOCALE_IDAYLZERO
    eMonthLeadingZero = LOCALE_IMONLZERO
    
    ePositiveSign = LOCALE_SPOSITIVESIGN
    eNegativeSign = LOCALE_SNEGATIVESIGN
    
    eTimeFormat = LOCALE_STIMEFORMAT
  End Enum

' the following functions call the local GetFormatString function to return
' the format string for the passed locale value
Public Function NegativeSign() As String
  
  Dim sNegativeSign$

  sNegativeSign = GetFormatString(eNegativeSign)
  
  If sNegativeSign = vbZLString Then sNegativeSign = "-"
  
  NegativeSign = sNegativeSign

End Function

Public Function PositiveSign() As String
  
  Dim sPositiveSign$

  sPositiveSign = GetFormatString(ePositiveSign)
  
  If sPositiveSign = vbZLString Then sPositiveSign = "+"
  
  PositiveSign = sPositiveSign

End Function

Public Function NumericDecimalSeparator() As String

  Dim sNumericSeparator$

  sNumericSeparator = GetFormatString(eNumericDecimalSeparator)
  
  If sNumericSeparator = vbZLString Then sNumericSeparator = "."
  
  NumericDecimalSeparator = sNumericSeparator

End Function

Public Function CurrencyDecimalSeparator() As String
  
  Dim sCurSeparator$

  sCurSeparator = GetFormatString(eCurrencyDecimalSeparator)
  
  If sCurSeparator = vbZLString Then sCurSeparator = "."
  
  CurrencyDecimalSeparator = sCurSeparator

End Function

Public Function CurrencyThousandsSeparator() As String

  Dim sThousandsSeparator$

  sThousandsSeparator = GetFormatString(eCurrencyThousandsSeparator)
  
  If sThousandsSeparator = vbZLString Then sThousandsSeparator = ","
  
  CurrencyThousandsSeparator = sThousandsSeparator
  
End Function

Public Function CurrencySymbol() As String
  
  Dim sCurSymbol$

  sCurSymbol = GetFormatString(eCurrencySymbol)
  
  If sCurSymbol = vbZLString Then sCurSymbol = "$"
  
  CurrencySymbol = sCurSymbol

End Function

Public Function IsCurrencySymbolSuffix() As Boolean
  ' on some systems, the currency is a suffix instead of a prefix for numbers.
  
  Dim sCurFormat$

  sCurFormat = GetFormatString(eCurrencyPrefixOrSuffixMode)
  
  Select Case sCurFormat
    Case "1", "3": IsCurrencySymbolSuffix = True
  End Select

End Function

Public Function CurrencySymbolFormatted() As String
  ' on some systems, the currency symbol has a space between it and the number
  
  Dim sCurSymbol$, sCurFormat$

  sCurSymbol = Me.CurrencySymbol
  
  sCurFormat = GetFormatString(eCurrencyPrefixOrSuffixMode)
  
  If sCurFormat = vbZLString Then sCurFormat = "0"
  
  Select Case sCurFormat
    Case "0", "1": CurrencySymbolFormatted = sCurSymbol
    Case "2":  CurrencySymbolFormatted = sCurSymbol & " "
    Case "3":  CurrencySymbolFormatted = " " & sCurSymbol
  End Select

End Function

Public Function ListSeparator() As String

  Dim sListSeparator$

  sListSeparator = GetFormatString(eListSeparator)
  
  If sListSeparator = vbZLString Then sListSeparator = ","
  
  ListSeparator = sListSeparator
  
End Function

Public Function DateSeparator() As String

  Dim sDateSeparator$

  sDateSeparator = GetFormatString(eDateSeparator)
  
  If sDateSeparator = vbZLString Then sDateSeparator = "/"
  
  DateSeparator = sDateSeparator
  
End Function

Public Function DayFormat() As String

  Dim sDayFormat$

  sDayFormat = GetFormatString(eDayLeadingZero)
  
  If sDayFormat = vbZLString Then sDayFormat = "0"
  
  Select Case sDayFormat
    Case "0": DayFormat = "d"
    Case "1": DayFormat = "dd"
  End Select

End Function

Public Function MonthFormat() As String

  Dim sMonthFormat$

  sMonthFormat = GetFormatString(eMonthLeadingZero)
  
  If sMonthFormat = vbZLString Then sMonthFormat = "0"
  
  Select Case sMonthFormat
    Case "0": MonthFormat = "m"
    Case "1": MonthFormat = "mm"
  End Select

End Function

Public Function FourDigitYearDateFormat() As String
  ' get the current short date format and add formatting for a 4 digit year.
  
  Dim sDateOrder$, sDateSeparator$, sDayFormat$, sMonthFormat$, nStrLen&, bError As Boolean

  With Me
    ' get the date separator
    sDateSeparator = .DateSeparator
  
    ' get the day format ("d" or "dd")
    sDayFormat = .DayFormat
  
    ' get the month format
    sMonthFormat = .MonthFormat
  End With
  
  sDateOrder = GetFormatString(eShortDateFormat)

  If sDateOrder = vbZLString Then sDateOrder = "0"

  ' return the correct format
  Select Case sDateOrder
    Case "0": FourDigitYearDateFormat = sMonthFormat & sDateSeparator & sDayFormat & sDateSeparator & "yyyy"
    Case "1": FourDigitYearDateFormat = sDayFormat & sDateSeparator & sMonthFormat & sDateSeparator & "yyyy"
    Case "2": FourDigitYearDateFormat = "yyyy" & sDateSeparator & sMonthFormat & sDateSeparator & sDayFormat
  End Select
  
End Function

Private Function GetFormatString(ByVal eFormat As LOCAL_FORMATS) As String
  ' this is the function where all of the work is done.
  
  Dim sFormatString$, nStrLen&

  ' get the length of the format string
  nStrLen = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, eFormat, sFormatString, API_FALSE)
  
  ' if the return wasn't zero...
  If nStrLen > 0 Then
    ' allocate string space and get the format string
    sFormatString = String$(nStrLen + 1, vbNullChar)
    nStrLen = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, eFormat, sFormatString, Len(sFormatString))
    
    ' if the return wasn't zero...
    If nStrLen > 0 Then
      ' strip any nulls and convert if necessarry
      sFormatString = StripNulls(sFormatString)
  
      ' check to see if the string is unicode
      If LenB(sFormatString) > 1 Then
        Dim aBytes(1) As Byte
      
        CopyMemory ByVal VarPtr(aBytes(0)), ByVal StrPtr(sFormatString), 2&
        
        If (aBytes(0) <> 0) And (aBytes(1) <> 0) Then sFormatString = StrConv(sFormatString, vbUnicode)
      
      Else
        sFormatString = StrConv(sFormatString, vbUnicode)
      End If
      
      ' return the string
      GetFormatString = sFormatString
    End If
  End If

End Function

Private Function StripNulls(ByVal sText As String) As String
  ' strips any nulls from the end of a string
  Dim nPosition&
  
  StripNulls = sText
  
  nPosition = InStr(sText, vbNullChar)
  If nPosition Then StripNulls = Left$(sText, nPosition - 1)
  If Len(sText) Then If Left$(sText, 1) = vbNullChar Then StripNulls = vbZLString
End Function

