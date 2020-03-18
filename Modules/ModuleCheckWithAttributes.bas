Attribute VB_Name = "ModuleCheckWithAttributes"
Sub Check(wb, p, r)
    ' Goal: Match new list with attribute list
    
    ' p: Attribute List
    ' r: List created by XML file
    
    ' Variables
    Dim q, cs As Object
    Dim i, j, k, m, n As Integer
    Dim s, t As String
    Dim dif As String
    Dim un, dt, eb, d, st As String
    
    ' These variables below represent columns. It would be enough to hardcode it but I chose to use the function "FindColumn" just in case there will be any changes
    ' made in the structure of the protocol.
    Dim id As Integer ' Identifier
    Dim at As Integer ' Beschreibung
    Dim tp As Integer ' Datentyp
    Dim se As Integer ' Standardeinheit
    Dim gr As Integer ' Gruppe
    Dim pf As Integer ' Pflichtfeld
    Dim ar As Integer ' Nur Artikel
    Dim di As Integer ' Dimension
    
    id = FindColumn(p, "Identifier", "Attribut-ID")
    at = FindColumn(p, "Beschreibung", "Attribut-Name")
    tp = FindColumn(p, "Typ", "Datentyp")
    se = FindColumn(p, "Standardeinheit", "Einheit physikalisch")
    gr = FindColumn(p, "Gruppe", "Gruppenzugehörigkeit")
    pf = FindColumn(p, "Pflichtfeld", "Pflichteintrag")
    ar = FindColumn(p, "Nur Artikel", "Artikel-/Produkdebene")
    di = FindColumn(p, "Dimension", "Dimension")
    
    ' There are 6 conditions that we can check to reach a 100& match and reference an attribute to an already existing ID
    ' 1. Attribute name
    ' 2. Datatype
    ' 3. Unit
    ' 4. Level
    ' 5. Dimension
    ' 6. Compliance
    
    ' For that I create an array with 6 booleans.
    ' If all are True, it will be a 100% match
    Dim b1(5) As Boolean
    Dim b2(5) As String
    Dim b As Boolean
    
    ' If there is a mismatch, my goal will be to show the attribute with the highest matches and document the mismatches in another column
    b2(0) = "Alles"
    b2(1) = "Datentyp"
    b2(2) = "Einheit"
    b2(3) = "Ebene"
    b2(4) = "Dimension"
    b2(5) = "Steuerung"
    
    
    ' 1. Go through all attributes in protocol
    i = 2
    Do Until r.Cells(i, 2) = ""
        s = r.Cells(i, 11)
        b1(0) = False
        b1(1) = False
        b1(2) = False
        b1(3) = False
        b1(4) = False
        b1(5) = False
        b = False
        ' Start value for counting number of attribute with highest matches, variable k will be reference
        k = 0
        j = 2
        Do Until b = True Or k = 6 Or p.Cells(j, id) = ""
            t = p.Cells(j, at)
            ' Check if (Compliance) is written as postfix of attribute. You don't see that in the XML so it has to be removed in the string variable
            If InStrRev(s, ")") <> Len(s) And InStrRev(t, ")") = Len(t) Then
                t = Left(t, InStrRev(t, "(") - 2)
            End If
            ' Check conditions
            b1(0) = (s = t)
            b1(1) = (InStrRev(p.Cells(j, tp), r.Cells(i, 7)) > 0)
            b1(2) = (r.Cells(i, 9) = p.Cells(j, se))
            b1(3) = ((r.Cells(i, 8) = "MerchandiseStyle" And p.Cells(j, ar) = "Nein") Or (r.Cells(i, 8) <> "MerchandiseStyle" And p.Cells(j, ar) = "Ja"))
            b1(4) = ((r.Cells(i, 10) = True And p.Cells(j, di) = "Ja") Or (r.Cells(i, 10) = False And p.Cells(j, di) = "Nein") Or (r.Cells(i, 10) = "" And p.Cells(j, ar) = "Nein"))
            b1(5) = ((r.Cells(i, 14) = "Ja" And p.Cells(j, gr) = "Verwaltungsattribute CoM") Or (r.Cells(i, 14) = "" And p.Cells(j, gr) <> "Verwaltungsattribute CoM"))
            ' Start value for counting matches for each loop but only if attribute name has been matched
            m = 0
            dif = ""
            ' Count number of matches
            If b1(0) = True Then
                m = m + 1
                For n = 1 To UBound(b1)
                    If b1(n) = True Then
                        m = m + 1
                    Else
                        If dif = "" Then
                            dif = b2(n)
                        Else
                            dif = dif & ", " & b2(n)
                        End If
                    End If
                Next
                ' Check if we have an attribute with a higher number of matches of its characteristics
                If m > k Then
                    k = m
                    r.Cells(i, 15) = dif
                    r.Cells(i, 16) = p.Cells(j, 1)
                    r.Cells(i, 6) = p.Cells(j, se)
                End If
            End If
            ' If k equals 6, then it is a 100% match
            ' The ID of the matching attribute will be inserted into column A
            If k = 6 Then
                b = True
                r.Cells(i, 1) = p.Cells(j, id)
                r.Cells(i, 15) = ""
                r.Cells(i, 16) = ""
            End If
            j = j + 1
        Loop
        ' If there is no 100% match then this attribute with its characteristics doesn't exist in database
        ' If the variable "dif" contains a string and is inserted into column O, the attribute exists by name but with different characteristics. Font of row will turn red.
        ' If dif does not contain any string then that attribute does not exist by name with different characteristics. In addition, font will turn bold.
        If r.Cells(i, 1) = "" Then
            Set q = Range(r.Cells(i, 1), r.Cells(i, 15))
            q.Font.Color = -16776961
            If r.Cells(i, 15) = "" Then q.Font.Bold = True
        End If
        ' Change name of characteristics as a preconditions for database import to match their IDs
        If r.Cells(i, 14) = "Ja" Then r.Cells(i, 12) = "Contentverwaltung"
        If r.Cells(i, 12) = "Maße & Gewicht" Then r.Cells(i, 12) = "Massangaben"
        
        i = i + 1
    Loop
    
    r.Cells.EntireColumn.AutoFit
    
    ' Create a Copy of protocol and delete all reds. This copy will be saved as a csv file and can be used for first import
    r.Copy after:=r
    Set cs = wb.ActiveSheet
    cs.Name = "PIM_Import"
    i = 1
    Do Until cs.Cells(i, 2) = ""
        If cs.Cells(i, 1) = "" Then
            cs.Rows(CStr(i) & ":" & CStr(i)).Delete
        Else
            i = i + 1
        End If
    Loop
    
    ' Delete columns that are not necessary
    i = i - 1
    cs.Range(cs.Cells(1, 5), cs.Cells(i, 17)).Delete
    
    ' If necessary, create attribute-ID with a specific logic and insert it in column A in the end.
    i = 2
    Do Until r.Cells(i, 2) = ""
        If r.Cells(i, 1) = "" Then
            ' First reference attribute name to string variable s.
            ' For the ID, some characters have to be replaced or removed like -
            s = r.Cells(i, 11)
            s = Replace(s, "-", "")
            un = r.Cells(i, 9)
            ' °  ==> Grad
            ' °C ==> GradC
            ' Next will be data type but not if data type refers to numbers (according to logic)
            ' Variable dt
            If r.Cells(i, 7) = "Wertemenge, mehrfach" Then
                dt = "_Wm"
            ElseIf r.Cells(i, 7) = "Wertemenge, einfach" Then
                dt = "_We"
            ElseIf InStrRev(r.Cells(i, 7), "Zeichenkette") > 0 Then
                dt = "_Zk"
            Else
                dt = ""
            End If
            s = s & dt
            
            ' Next add unit but replace or remove some characters that are not valid for ID
            If un <> "" Then
                un = Replace(un, "°", "Grad")
                un = Replace(un, "²", "2")
                un = Replace(un, "³", "3")
                un = Replace(un, "%", "Prozent")
                If InStrRev(un, "B") > 1 Then un = Replace(un, "B", "b")
                'un = Replace(un, "Std", "h")
                'un = Replace(un, "Stunden", "h")
                'un = Replace(un, "Stunde", "h")
                un = Replace(un, "Quadratmeter", "m2")
                un = Replace(un, "Karat", "ct")
                un = Replace(un, "/", "pro")
                un = Replace(un, "Kilogramm", "kg")
                un = Replace(un, "kilogramm", "kg")
                un = Replace(un, "Kilometer", "km")
                un = Replace(un, "kilometer", "km")
                un = Replace(un, "Kilowatt", "kW")
                un = Replace(un, "kilowatt", "km")
                un = Replace(un, "-", "")
                un = Replace(un, "Kilowattstunde", "kWh")
                un = Replace(un, "Watt", "W")
                un = Replace(un, "Kilowatt", "kW")
                un = Replace(un, "Kubikmeter", "m3")
                un = Replace(un, "Minute", "min")
                un = Replace(un, "Minuten", "min")
                un = Replace(un, "Liter", "l")
                un = Replace(un, "Sekunde", "s")
                un = Replace(un, "Milli", "m")
                un = Replace(un, "Pixel", "px")
                un = Replace(un, "Stück", "Stk")
                un = Replace(un, ".", "")
                un = Replace(un, "Tag", "Tage")
                un = Replace(un, "Tag(e)", "Tage")
                un = Replace(un, ChrW(937), "Ohm")
                un = Replace(un, "meter", "m")
                un = Replace(un, "·", "")
                un = Replace(un, """", "")
                un = Replace(un, "·", "")
                
                If un <> "" Then s = s & "_" & un
            End If
            ' Since we deal with German words, we need to replace further letters
            s = Replace(s, "Ä", "ae")
            s = Replace(s, "Ö", "oe")
            s = Replace(s, "Ü", "ue")
            s = Replace(s, "ä", "ae")
            s = Replace(s, "ö", "oe")
            s = Replace(s, "ü", "ue")
            s = Replace(s, "ß", "ss")
            s = Replace(s, " ", "")
            s = Replace(s, "/", "")
            s = Replace(s, "(", "")
            s = Replace(s, ")", "")
            
            
            ' Product or Article level
            ' Variable eb
            If r.Cells(i, 8) = "MerchandiseStyle" Then
                eb = "_Produkt"
            Else
                eb = "_Artikel"
            End If
            s = s & eb
            
            ' Dimensioned?
            ' Variable d
            If r.Cells(i, 10) = True Then
                d = "_DIM"
            Else
                d = ""
            End If
            s = s & d
            
            ' Compliance?
            ' Variable st
            st = ""
            If r.Cells(i, 14) = "Ja" Then st = "_Steuerung"
            s = s & st
            
            ' Now we can insert
            r.Cells(i, 1) = s
        End If
        i = i + 1
    Loop
End Sub


