Attribute VB_Name = "ModuleProduktArtikel"
Sub prodart(q, r)
    ' Goal: Check if attribute is dimensioned. Only item based attributes
    
    ' q: File/Sheet with dimensioned attributes
    ' r: New created protocol
    
    Dim i, j As Integer
    Dim b As Boolean
    Dim s, t As String
    
    ' Go through all attributes in protocol
    i = 2
    Do Until r.Cells(i, 2) = ""
        If r.Cells(i, 8) <> "MerchandiseStyle" Then
            s = r.Cells(i, 11)
            b = False
            j = 2
            ' Check for match with dimensioned attributes
            Do Until q.Cells(j, 1) = "" Or b = True
                t = q.Cells(j, 1)
                If s = t Then
                    b = True
                Else
                    j = j + 1
                End If
            Loop
        r.Cells(i, 10) = b
        End If
        i = i + 1
    Loop
End Sub

