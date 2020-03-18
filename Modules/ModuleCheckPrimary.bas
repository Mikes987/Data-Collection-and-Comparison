Attribute VB_Name = "ModuleCheckPrimary"
Sub CheckPrimary(o, p)
    ' Goal: Compare Primary Data Set and protocol
    
    ' o: Protocol
    ' p: Primary Dataset
    
    Dim i, j As Integer
    Dim k, m As Integer
    Dim s, t, u As String
    Dim b As Boolean
    Dim b1(3) As Boolean
    Dim b2(3) As String
    
    ' Variables below reference header/columns
    Dim at As Integer ' Attribute name
    Dim eb As Integer ' Attribut level
    Dim pf As Integer ' Attribut mandatory option
    Dim dt As Integer ' Attribut datatype
    Dim un As Integer ' Attribut unit
    
    at = FindColumn(o, "Attribut", "Attribut, Beschreibung")
    eb = FindColumn(o, "xFit Ebene", "Ebene")
    pf = FindColumn(o, "Pflichttyp", "Pflichtangabe")
    dt = FindColumn(o, "xFit Datentyp", "Datentyp")
    un = FindColumn(o, "xFit Einheit", "Einheit, physikalisch")
    
    ' What can we check and compare?
    b2(0) = "Datentyp"       ' datatype
    b2(1) = "Ebene"          ' level
    b2(2) = "Pflichteintrag" ' Mandatory option
    b2(3) = "Einheit"        ' Unit
    
    ' Loop 1: Go through all attributes in protocol
    i = 2
    k = 0
    Do Until o.Cells(i, at) = ""
        s = o.Cells(i, at)
        j = 1
        b = False
        ' String variable below collects mismatches
        u = ""
        ' Loop 2: Go through Primary data file in that column (so far column 6) that contains attribute names. Search for match.
        Do Until b = True Or (p.Cells(j, 6) = "" And p.Cells(j, 6).MergeCells = False)
            t = p.Cells(j, 6)
            If s = t Then
                b = True
                ' Set all values of boolean array to False
                For k = 0 To UBound(b1)
                    b1(k) = False
                Next
                ' Match conditions
                b1(0) = (o.Cells(i, dt) = "Zeichenkette" And p.Cells(j, 4) = "Z") Or (o.Cells(i, dt) = "Zahl, dezimal" And p.Cells(j, 4) = "D") Or _
                        (o.Cells(i, dt) = "Zahl, ganzzahlig" And p.Cells(j, 4) = "G") Or (o.Cells(i, dt) = "Wertemenge, einfach" And p.Cells(j, 4) = "") Or _
                        (o.Cells(i, dt) = "Wertemenge, mehrfach" And p.Cells(j, 4) = "" And p.Cells(j + 1, 6) = "")
                b1(1) = (o.Cells(i, eb) = "MerchandiseStyle" And p.Cells(j, 2) = "P") Or (InStr(o.Cells(i, eb), "Item") > 0 And p.Cells(j, 2) = "A" Or (InStr(o.Cells(i, eb), "Item") > 0 And p.Cells(j, 2) = "V"))
                b1(2) = (o.Cells(i, pf) = "Pflicht" And p.Cells(j, 3) = "X") Or (o.Cells(i, pf) = "Optional" And p.Cells(j, 3) = "")
                b1(3) = (o.Cells(i, un) = p.Cells(j, 7) Or o.Cells(i, un) = "Zoll" And p.Cells(j, 7) = """")
                ' Count matches
                For k = 0 To UBound(b1)
                    If b1(k) = False Then
                        ' If there are missmatches, color the interior cells to make them visible
                        u = u & b2(k) & ","
                        o.Range(o.Cells(i, 1), o.Cells(i, 17)).Interior.Color = 65535
                    End If
                Next
            Else
                j = j + 1
            End If
        Loop
        
        ' If u is not empty, set it into the most right column in the specific line
        If u <> "" Then
            u = Left(u, Len(u) - 1)
            o.Cells(i, 17) = u
        End If
        ' The variable b tells us if that attribute exists in the XML file but not in Primary Data Set
        ' Then it actually shouldn't exist in database
        If b = False Then
            o.Range(o.Cells(i, 1), o.Cells(i, 17)).Interior.Color = 13431551
            o.Cells(i, 17) = "Attribut nicht in Primary Live"
        End If
        
        i = i + 1
    Loop
End Sub
