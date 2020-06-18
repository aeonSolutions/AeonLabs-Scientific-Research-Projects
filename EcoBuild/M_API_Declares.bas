Attribute VB_Name = "M_API_Declares"
  Option Explicit
  ' demo project showing how to use the GetLocaleInfo API function
  ' by Bryan Stafford of New Vision Software® - newvision@mvps.org
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.

  Public Const vbZLString As String = ""
  Public Const API_FALSE As Long = &H0&
  Public Const API_TRUE As Long = &H1&

  Public Const LOCALE_SYSTEM_DEFAULT As Long = &H800
  Public Const LOCALE_USER_DEFAULT As Long = &H400

  Public Const LOCALE_SLIST As Long = &HC         '  list item separator
  Public Const LOCALE_IMEASURE As Long = &HD         '  0 as long = metric, 1 as long = US
  
  Public Const LOCALE_SDECIMAL As Long = &HE         '  decimal separator
  Public Const LOCALE_STHOUSAND As Long = &HF         '  thousand separator
  Public Const LOCALE_SGROUPING As Long = &H10        '  digit grouping
  Public Const LOCALE_IDIGITS As Long = &H11        '  number of fractional digits
  Public Const LOCALE_ILZERO As Long = &H12        '  leading zeros for decimal
  Public Const LOCALE_SNATIVEDIGITS As Long = &H13        '  native ascii 0-9
  
  Public Const LOCALE_SCURRENCY As Long = &H14        '  local monetary symbol
  Public Const LOCALE_SINTLSYMBOL As Long = &H15        '  intl monetary symbol
  Public Const LOCALE_SMONDECIMALSEP As Long = &H16        '  monetary decimal separator
  Public Const LOCALE_SMONTHOUSANDSEP As Long = &H17        '  monetary thousand separator
  Public Const LOCALE_SMONGROUPING As Long = &H18        '  monetary grouping
  Public Const LOCALE_ICURRDIGITS As Long = &H19        '  # local monetary digits
  Public Const LOCALE_IINTLCURRDIGITS As Long = &H1A        '  # intl monetary digits
  Public Const LOCALE_ICURRENCY As Long = &H1B        '  positive currency mode
  Public Const LOCALE_INEGCURR As Long = &H1C        '  negative currency mode
  
  Public Const LOCALE_SDATE As Long = &H1D        '  date separator
  Public Const LOCALE_STIME As Long = &H1E        '  time separator
  Public Const LOCALE_SSHORTDATE As Long = &H1F        '  short date format string
  Public Const LOCALE_SLONGDATE As Long = &H20        '  long date format string
  Public Const LOCALE_STIMEFORMAT As Long = &H1003      '  time format string
  Public Const LOCALE_IDATE As Long = &H21        '  short date format ordering
  Public Const LOCALE_ILDATE As Long = &H22        '  long date format ordering
  Public Const LOCALE_ITIME As Long = &H23        '  time format specifier
  Public Const LOCALE_ICENTURY As Long = &H24        '  century format specifier
  Public Const LOCALE_ITLZERO As Long = &H25        '  leading zeros in time field
  Public Const LOCALE_IDAYLZERO As Long = &H26        '  leading zeros in day field
  Public Const LOCALE_IMONLZERO As Long = &H27        '  leading zeros in month field
  Public Const LOCALE_S1159 As Long = &H28        '  AM designator
  Public Const LOCALE_S2359 As Long = &H29        '  PM designator
  
  Public Const LOCALE_SDAYNAME1 As Long = &H2A        '  long name for Monday
  Public Const LOCALE_SDAYNAME2 As Long = &H2B        '  long name for Tuesday
  Public Const LOCALE_SDAYNAME3 As Long = &H2C        '  long name for Wednesday
  Public Const LOCALE_SDAYNAME4 As Long = &H2D        '  long name for Thursday
  Public Const LOCALE_SDAYNAME5 As Long = &H2E        '  long name for Friday
  Public Const LOCALE_SDAYNAME6 As Long = &H2F        '  long name for Saturday
  Public Const LOCALE_SDAYNAME7 As Long = &H30        '  long name for Sunday
  Public Const LOCALE_SABBREVDAYNAME1 As Long = &H31        '  abbreviated name for Monday
  Public Const LOCALE_SABBREVDAYNAME2 As Long = &H32        '  abbreviated name for Tuesday
  Public Const LOCALE_SABBREVDAYNAME3 As Long = &H33        '  abbreviated name for Wednesday
  Public Const LOCALE_SABBREVDAYNAME4 As Long = &H34        '  abbreviated name for Thursday
  Public Const LOCALE_SABBREVDAYNAME5 As Long = &H35        '  abbreviated name for Friday
  Public Const LOCALE_SABBREVDAYNAME6 As Long = &H36        '  abbreviated name for Saturday
  Public Const LOCALE_SABBREVDAYNAME7 As Long = &H37        '  abbreviated name for Sunday
  Public Const LOCALE_SMONTHNAME1 As Long = &H38        '  long name for January
  Public Const LOCALE_SMONTHNAME2 As Long = &H39        '  long name for February
  Public Const LOCALE_SMONTHNAME3 As Long = &H3A        '  long name for March
  Public Const LOCALE_SMONTHNAME4 As Long = &H3B        '  long name for April
  Public Const LOCALE_SMONTHNAME5 As Long = &H3C        '  long name for May
  Public Const LOCALE_SMONTHNAME6 As Long = &H3D        '  long name for June
  Public Const LOCALE_SMONTHNAME7 As Long = &H3E        '  long name for July
  Public Const LOCALE_SMONTHNAME8 As Long = &H3F        '  long name for August
  Public Const LOCALE_SMONTHNAME9 As Long = &H40        '  long name for September
  Public Const LOCALE_SMONTHNAME10 As Long = &H41        '  long name for October
  Public Const LOCALE_SMONTHNAME11 As Long = &H42        '  long name for November
  Public Const LOCALE_SMONTHNAME12 As Long = &H43        '  long name for December
  Public Const LOCALE_SABBREVMONTHNAME1 As Long = &H44        '  abbreviated name for January
  Public Const LOCALE_SABBREVMONTHNAME2 As Long = &H45        '  abbreviated name for February
  Public Const LOCALE_SABBREVMONTHNAME3 As Long = &H46        '  abbreviated name for March
  Public Const LOCALE_SABBREVMONTHNAME4 As Long = &H47        '  abbreviated name for April
  Public Const LOCALE_SABBREVMONTHNAME5 As Long = &H48        '  abbreviated name for May
  Public Const LOCALE_SABBREVMONTHNAME6 As Long = &H49        '  abbreviated name for June
  Public Const LOCALE_SABBREVMONTHNAME7 As Long = &H4A        '  abbreviated name for July
  Public Const LOCALE_SABBREVMONTHNAME8 As Long = &H4B        '  abbreviated name for August
  Public Const LOCALE_SABBREVMONTHNAME9 As Long = &H4C        '  abbreviated name for September
  Public Const LOCALE_SABBREVMONTHNAME10 As Long = &H4D        '  abbreviated name for October
  Public Const LOCALE_SABBREVMONTHNAME11 As Long = &H4E        '  abbreviated name for November
  Public Const LOCALE_SABBREVMONTHNAME12 As Long = &H4F        '  abbreviated name for December
  Public Const LOCALE_SABBREVMONTHNAME13 As Long = &H100F
  
  Public Const LOCALE_SPOSITIVESIGN As Long = &H50        '  positive sign
  Public Const LOCALE_SNEGATIVESIGN As Long = &H51        '  negative sign
  Public Const LOCALE_IPOSSIGNPOSN As Long = &H52        '  positive sign position
  Public Const LOCALE_INEGSIGNPOSN As Long = &H53        '  negative sign position
  Public Const LOCALE_IPOSSYMPRECEDES As Long = &H54        '  mon sym precedes pos amt
  Public Const LOCALE_IPOSSEPBYSPACE As Long = &H55        '  mon sym sep by space from pos amt
  
  Public Const LOCALE_INEGSYMPRECEDES As Long = &H56        '  mon sym precedes neg amt
  Public Const LOCALE_INEGSEPBYSPACE As Long = &H57        '  mon sym sep by space from neg amt


  Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale&, ByVal LCType&, ByVal lpLCData$, ByVal cchData&) As Long

  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

