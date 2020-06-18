Attribute VB_Name = "Gobal_vars"
Public fMainForm As frmMain
Public check_db As Boolean

'Specify the files that will appear in the dialog open/save box'
Public Const dialogs_filter = "All Files (*.*)|*.*|EcoBuild Files" & "(*.eco)|*.eco"
'extension used in the save/open file
Public Const filename_extension = "eco"
'build date of the program
Public Const build_date = "19-01-2005"

Type impact
    energy As Single
    co2 As Single
    so2 As Single
    nox As Single
    water As Single
    costs As Single
End Type

 Type default_type
    energy As Single
    co2 As Single
    so2 As Single
    nox As Single
    water As Single
End Type

 Type concrete_type
    name As String
    date As String
    description As String
    wood As default_type
    steel As default_type
    cement As default_type
    agregates As default_type
    water As default_type
End Type

 Type trans
    distance As Single
    co2 As Single
    so2 As Single
    nox As Single
End Type

 Type steel_type
    name As String
    date As String
    description As String
    steel As default_type
    transport As trans
End Type
Type impact_values
    energy(1 To 1, 1 To 2) As Double
    co2(1 To 1, 1 To 2) As Double
    nox(1 To 1, 1 To 2) As Double
    so2(1 To 1, 1 To 2) As Double
    water(1 To 1, 1 To 2) As Double
    costs(1 To 1, 1 To 2) As Double
    total(1 To 1, 1 To 2) As Double
End Type
Type document_proprieties
    impact_metal As impact
    impact_concrete As impact
    impact_transport As impact
    impact_total As impact
    volume_concrete As Double
    volume_wood As Double
    total_weight As Double
    cement_qty As Double
    aggregates_qty As Double
    armour_qty As Double
    database As String
    cement As String
    aggregates As String
    metal_cost As String
    concrete_cost As String
    dados As impact_values
End Type

Public concrete() As concrete_type
Public metalic() As steel_type

Public doc_props() As document_proprieties
