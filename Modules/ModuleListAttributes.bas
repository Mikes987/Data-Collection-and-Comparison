Attribute VB_Name = "ModuleAttributeListen"
Sub List(a, n, o)
    ' Goal: Insert values of array from "DoSearch" into a new sheet in a specific order to fit specification of database
    
    ' Variables
    Dim k(16) As String
    Dim i As Integer
    
    ' Array for headers of columns
    k(0) = "Attribut_identifier"
    k(1) = "Root_Kategorie"
    k(2) = "Kategorie"
    k(3) = "Pflichttyp"
    k(4) = "Standardwert"
    k(5) = "Einheit"
    k(6) = "xFit Datentyp"
    k(7) = "xFit Ebene"
    k(8) = "xFit Einheit"
    k(9) = "Dimensioniert"
    k(10) = "Attribut"
    k(11) = "Gruppe"
    k(12) = "Einheit, ausgeschrieben"
    k(13) = "Steuerung"
    k(14) = "Unterschied in PIM"
    k(15) = "ID mit höchstem Match"
    k(16) = "Unterschied zu Primary"
    

    'Fill columns
    For i = 0 To UBound(k)
        o.Cells(1, i + 1) = k(i)
    Next
    o.Rows("1:1").Font.Bold = True
    
    ' Fill attributes and characteristics of arrays into the specific columns
    For i = 0 To UBound(a, 2)
        o.Cells(i + 2, 2) = "VERSION"
        o.Cells(i + 2, 3) = "PBK_" & n
        o.Cells(i + 2, 4) = a(2, i)
        o.Cells(i + 2, 7) = a(1, i)
        o.Cells(i + 2, 8) = a(3, i)
        o.Cells(i + 2, 9) = a(4, i)
        o.Cells(i + 2, 12) = a(5, i)
        o.Cells(i + 2, 11) = a(0, i)
        o.Cells(i + 2, 13) = a(6, i)
        If a(7, i) = "0" Then o.Cells(i + 2, 14) = "Ja"
    Next
    o.Cells.EntireColumn.AutoFit

    ' Sort according to attribute names
    i = i + 1
    o.Range(o.Cells(2, 1), o.Cells(i, 14)).Sort key1:=o.Range("K2"), Header:=xlNo
End Sub
