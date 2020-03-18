Attribute VB_Name = "ModulePrepareImport"
Sub PrepareImport(o)
    ' Goal: If a new import file will be created set titles for columns and place them into specific columns
    
    Dim a(15) As String
    Dim i As Integer
    
    
    a(0) = "Gruppe"
    a(1) = "Attribut"
    a(2) = "Wert"
    a(3) = "Identifier"
    a(4) = "Beschreibung"
    a(5) = "Sequenz"
    a(6) = "Dimension"
    a(7) = "Vererbung"
    a(8) = "Nur Artikel"
    a(9) = "Nur Farbebene"
    a(10) = "Typ"
    a(11) = "Einheit"
    a(12) = "Standardeinheit"
    a(13) = "Pflichtfeld"
    a(14) = "Benutzerrecht"
    a(15) = "Kommentar"
    
    For i = 0 To UBound(a)
        With o.Cells(1, i + 1)
            .Value = a(i)
            .Font.Bold = True
        End With
    Next
    
    o.Range("G1:P1").Interior.Color = 3243501
End Sub

