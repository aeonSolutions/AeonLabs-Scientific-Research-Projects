Attribute VB_Name = "Engine"
Option Explicit 'Transformas all declared variables into global variables

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
'this func checks the existence of the first concrete/metallic entry
'and returns true if there's an entry in the document(doc).lista
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
Public Function enabler(where As String, what As String) As String
Dim doc As Integer
Dim i As Integer

enabler = "Null"
doc = current_form
If FState(doc).count = 0 Then
    Exit Function
End If
For i = 1 To FState(doc).count
    With document(doc)
        .lista.Row = i
        .lista.Col = 0
        If (.lista.Text = "Conc.Beam" Or .lista.Text = "Conc.Pillar") And where = "Concrete" Then
            If what = "Database" Then
                .lista.Col = 18
                enabler = .lista.Text
                Exit Function
            End If
            If what = "Cement" Then
                .lista.Col = 15
                enabler = .lista.Text
                Exit Function
            End If
            If what = "Aggregates" Then
                .lista.Col = 16
                enabler = .lista.Text
                Exit Function
            End If
            If what = "Costs" Then
                .lista.Col = 17
                enabler = .lista.Text
                Exit Function
            End If
        End If
        If (.lista.Text = "Met.Beam" Or .lista.Text = "Met.Pillar") And where = "Metallic" Then
            If what = "Database" Then
                .lista.Col = 18
                enabler = .lista.Text
                Exit Function
            End If
            If what = "Costs" Then
                .lista.Col = 17
                enabler = .lista.Text
                Exit Function
            End If
        End If
    End With
Next i
End Function

'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°
' this sub loads the values in the selected database
'for the doc document
'¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°`°º¤ø,¸¸,ø¤º°`°º¤¤º°

Public Sub load_database()

Dim i As Integer
Dim filename As String
Dim chain As String
Dim r() As String
Dim s() As String
Dim num As Integer

    ' loading concrete database
    ReDim concrete(1)
    Err.Clear
    On Error Resume Next
    filename = App.path & "\database\concrete.dbs"
    Open filename For Input As #1
    If Err.Number = 0 Then ' file not found!?
        Input #1, num
        ReDim concrete(num + 1)
        i = 0
        While Not EOF(1)
            Input #1, chain
            i = i + 1
            r() = Split(chain, "@")
            s() = Split(r(0), "#")
            With concrete(i)
                .name = s(0)
                .date = s(1)
                .description = s(2)
                s() = Split(r(1), "#")
                With .wood
                    .co2 = str2str(s(0))
                    .energy = str2str(s(1))
                    .nox = str2str(s(2))
                    .so2 = str2str(s(3))
                    .water = str2str(s(4))
                End With
                s() = Split(r(2), "#")
                With .cement
                    .co2 = str2str(s(0))
                    .energy = str2str(s(1))
                    .nox = str2str(s(2))
                    .so2 = str2str(s(3))
                    .water = str2str(s(4))
                End With
                s() = Split(r(3), "#")
                With .steel
                    .co2 = str2str(s(0))
                    .energy = str2str(s(1))
                    .nox = str2str(s(2))
                    .so2 = str2str(s(3))
                    .water = str2str(s(4))
                End With
                s() = Split(r(4), "#")
                With .agregates
                    .co2 = str2str(s(0))
                    .energy = str2str(s(1))
                    .nox = str2str(s(2))
                    .so2 = str2str(s(3))
                    .water = str2str(s(4))
                End With
                s() = Split(r(5), "#")
                With .water
                    .co2 = str2str(s(0))
                    .energy = str2str(s(1))
                    .nox = str2str(s(2))
                    .so2 = str2str(s(3))
                    .water = str2str(s(4))
                End With
            End With
        Wend
        
    End If
    Close #1
End Sub

