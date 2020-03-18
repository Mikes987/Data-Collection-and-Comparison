Attribute VB_Name = "ModuleDoSearch"
Function DoSearch(o, p, at, mv, pf, tp, va, td, gr, st, uo, un)
    ' Goal: Create a 2d-array that collects all information from the XML file
    
    ' What are all the variables in the brackets?
    ' o:   XML-file as parsed in excel
    ' p:   PIM-PRIMARY-Attribute
    ' at:  Column with attribute names
    ' mv:  Column with information about multi or single default value option, True for multi, False for single
    ' pf:  Column with information about Mandatory attribute? True for yes, False for no
    ' tp:  Column describing data type
    ' va:  Column with default values if they exist
    ' td:  Column describing product or article level
    ' gr:  Column describing product group
    ' st:  Column describing compliance (yes if value is 0)
    ' uo:  Column with unit written in full (kilogramm)
    ' un:  Column with unit as physical notation (kg)
    
    ' Variables
    Dim q As Object
    Dim i, j, k As Long
    Dim s As String
    
    ' String variables that store content of parsed XML file per row before inserting into array
    Dim s0, s1, s2, s3, s4, s5, s6, s7 As String
    
    ' s0: Attribute name
    ' s1: Datatype
    ' s2: Mandatory option
    ' s3: Item, ItemOption oder MerchandiseStyle
    ' s4: Unit physical
    ' s5: Product group
    ' s6: Unit written in full
    ' s7: Compliance
    
    Dim b As Boolean
    Dim fi(14) As String ' Filterarray
    Dim a() As String    ' Multidimensional array to store attribute and its characteristics
    
    ' Filterarray - Not all attributes are necessary to be stored and can be skipped
    fi(0) = "Gratiskurztext"
    fi(1) = "Gratislangtext"
    fi(2) = "Produkt -Name"
    fi(3) = "Produkttyp"
    fi(4) = "Selling Point 1"
    fi(5) = "Selling Point 2"
    fi(6) = "Selling Point 3"
    fi(7) = "Selling Point 4"
    fi(8) = "Selling Point 5"
    fi(9) = "Serie"
    fi(10) = "Serie Technischer Name"
    fi(11) = "Set-Typ"
    fi(12) = "Grundfarbe"
    fi(13) = "Set-Info"
    fi(14) = "Produkt-Name"
    
    ' Content of parsed XML File begins at row 2. First we create an until-loop and go through the entire list.
    i = 2
    Do Until o.Cells(i, 1) = ""
        ' If the loop begins, the array is not dimensioned or is empty respectively and has to be filled first.
        ' Therefore, we write an if-then-else sequence and if the array has no content, it will go through the "else" first.
        If Not Not a Then
            ' Now the array is dimensioned it the loop will always go through "then"
            s0 = o.Cells(i, at)
            b = False
            For j = 0 To UBound(fi)
                If s0 = fi(j) Or InStrRev(s0, "Packstück") > 0 Then
                    b = True
                    Exit For
                End If
            Next
            ' Ein Attribut kann beispielsweise sowohl als Zahl als auch als Wertemenge deklariert sein, daher kategorisieren wir den Array zuerst und prüfen dann erst,
            ' ob dieser schon vorhanden ist.
            If b = False Then
                s1 = o.Cells(i, tp)
                s3 = o.Cells(i, td)
                s4 = o.Cells(i, un)
                s4 = Replace(s4, "·", "")
                s4 = Replace(s4, """", "Zoll")
                s5 = o.Cells(i, gr)
                s6 = o.Cells(i, uo)
                s7 = o.Cells(i, st)
                If o.Cells(i, va) <> "" And o.Cells(i, mv) = False Then
                    s1 = "Wertemenge, einfach"
                    'If s4 <> "V" Then s4 = ""
                ElseIf o.Cells(i, va) <> "" And o.Cells(i, mv) = True Then
                    s1 = "Wertemenge, mehrfach"
                    'If s4 <> "V" Then s4 = ""
                End If
                If o.Cells(i, pf) = True Then
                    s2 = "Pflicht"
                Else
                    s2 = "Optional"
                End If
                ' Wir prüfen nun, ob diese Einträge so schon im Array existieren und ignorieren gegebenenfalls die neue Zusammenstellung.
                For j = 0 To UBound(a, 2)
                    If s0 = a(0, j) And s1 = a(1, j) And s2 = a(2, j) And s3 = a(3, j) And s4 = a(4, j) And s5 = a(5, j) And s6 = a(6, j) And s7 = a(7, j) Then
                        b = True
                        Exit For
                    End If
                Next
                ' Existierte dieser Eintrag nicht, dann können wir ihn im Array ergänzen.
                If b = False Then
                    ReDim Preserve a(7, j)
                    a(0, j) = s0
                    a(1, j) = s1
                    a(2, j) = s2
                    a(3, j) = s3
                    a(4, j) = s4
                    a(5, j) = s5
                    a(6, j) = s6
                    a(7, j) = s7
                End If
            End If
        Else
            ' We save the attribute name and check for match with filterarray
            s0 = o.Cells(i, at)
            b = False
            For j = 0 To UBound(fi)
                If s0 = fi(j) Or InStrRev(s0, "Packstück") > 0 Then
                    b = True
                    Exit For
                End If
            Next
            If b = False Then
                ' If the attribute does not need to be filtered, we store all its characteristics into string variables
                s1 = o.Cells(i, tp)
                s3 = o.Cells(i, td)
                s4 = o.Cells(i, un)
                ' The database cannot read certain characters in its unit content, they have to be removed or replaced
                s4 = Replace(s4, "·", "")
                s4 = Replace(s4, """", "Zoll")
                s5 = o.Cells(i, gr)
                s6 = o.Cells(i, uo)
                s7 = o.Cells(i, st)
                ' Does it contain default values and is it to be set as single or multiple options?
                If o.Cells(i, va) <> "" And o.Cells(i, mv) = False Then
                    s1 = "Wertemenge, einfach"
                ElseIf o.Cells(i, va) <> "" And o.Cells(i, mv) = True Then
                    s1 = "Wertemenge, mehrfach"
                End If
                ' Mandatory or optional attribute
                If o.Cells(i, pf) = True Then
                    s2 = "Pflicht"
                Else
                    s2 = "Optional"
                End If
                ' All characteristics are handled to fit defaults of database and will be stored in array
                ReDim a(7, 0)
                a(0, 0) = s0
                a(1, 0) = s1
                a(2, 0) = s2
                a(3, 0) = s3
                a(4, 0) = s4
                a(5, 0) = s5
                a(6, 0) = s6
                a(7, 0) = s7
            End If
        End If
        i = i + 1
    Loop
    
    ' Change names of data types to fit defaults of database
    For i = 0 To UBound(a, 2)
        If a(1, i) = "ZEICHENKETTE" Then a(1, i) = "Zeichenkette"
        If a(1, i) = "DEZIMAL" Then a(1, i) = "Zahl, dezimal"
        If a(1, i) = "GANZZAHL" Then a(1, i) = "Zahl, ganzzahlig"
        If a(1, i) = "FLIESSKOMMA" Then a(1, i) = "Zahl, dezimal"
    Next
    
    ' Done, array can be returned
    DoSearch = a
End Function
