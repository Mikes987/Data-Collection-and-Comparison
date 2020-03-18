Attribute VB_Name = "ModuleImport"
Sub Import(o, p, q)
    ' Goal: Store attributes, characteristics and default values into an import file in a specific order.
    
    ' o: Protocol made by XML file
    ' p: Default values
    ' q: (Newly) created import file and sheet
    
    ' Variables
    Dim i, j, l, m, n As Integer
    Dim k As Long
    Dim di, ar, tp, ui, un, va As Integer
    Dim b, b1 As Boolean
    Dim pbk, uns As String
    Dim r, s, t As String
    Dim z As Long
    
    ' General index for rows, set as long datatype just in case.
    z = 1
    
    ' At first, we need to address certain columns
    ' di: Column with header "Dimension"
    ' ar: Column with header "Nur Artikel"
    ' tp: Column with header "Typ"
    ' ui: Column with header "Einheit"
    ' un: Column with header "Standardeinheit"
    
    ' If a new product is inserted, these columns have to be moved one cell to the right because a new column will be inserted
    di = FindColumn(q, "Dimension", "Dimension")
    ar = FindColumn(q, "Nur Artikel", "Artikel-/Produktebene")
    tp = FindColumn(q, "Typ", "Datentyp")
    ui = FindColumn(q, "Einheit", "Einheit-ID")
    un = FindColumn(q, "Standardeinheit", "Einheit, ausgeschrieben")
    
    ' ID directory in database for units
    uns = "BMD_UOMS"
    
    ' We always pretend that this new import file may be filled with a product and its attributes and default values
    pbk = o.Cells(2, 3) ' product name, will always be stored in row 1
    j = 1
    b = False
    Do Until b = True Or j = di
        If q.Cells(1, j) = pbk Then
            b = True
        Else
            j = j + 1
        End If
    Loop
    ' If the product name does not exist, create a new column to the left of "Dimension" and place the name in there.
    If b = False Then
        r = Split(q.Cells(1, di).Address, "$")(1)
        q.Columns(r & ":" & r).Insert
        q.Cells(1, di) = pbk
        di = di + 1
        ar = ar + 1
        tp = tp + 1
        ui = ui + 1
        un = un + 1
    End If
    
    ' Attributes and default values will  be inserted. We can loop the list by knowing that either column A, B or C contain values until the list ends.
    ' Loop 1: Go through all attributes in protocol
    i = 2
    Do Until o.Cells(i, 2) = ""
        k = 3
        b = False
        ' Loop 2: Check if attribute ID already is inserted into the import file through another product
        Do Until (q.Cells(k, 1) = "" And q.Cells(k, 2) = "" And q.Cells(k, 3) = "") Or b = True
            ' Test 1: Check if that row contains an attribute ID
            If q.Cells(k, 1) <> "" And q.Cells(k, 2) <> "" Then
                ' Test 2: Check for match between attribute ID in protocol and import file
                If o.Cells(i, 1) = q.Cells(k, 4) Then
                    z = z + 1
                    b = True
                    q.Cells(k, j) = "x"
                    ' Test 3: Check if this attribute contains default values. If yes, then these default values have to be checked for matches as well.
                    If InStrRev(q.Cells(k, tp), "Wertemenge") > 0 Then
                        z = z + 1
                        ' Test 4: Check default values. For that, find the column in sheet for default values in protocol file. Row 1 contains the attribute IDs
                        m = FindColumn(p, o.Cells(i, 1), o.Cells(i, 1))
                        n = 2
                        b1 = False
                        ' Loop 3: Go through all default values in sheet "default values" and check if they already exist in import file
                        Do Until p.Cells(n, m) = ""
                            va = k + 1
                            b1 = False
                            ' Loop 4: Go through inserted default values in
                            Do Until b1 = True Or q.Cells(va, 3) = ""
                                ' Prüfung 5: Check default values for a match
                                If q.Cells(va, 5) = p.Cells(n, m) Then
                                    b1 = True
                                    z = z + 1
                                Else
                                    va = va + 1
                                End If
                            Loop
                            ' If the boolean variable stays "False" then we have to insert a new default value.
                            If b1 = False Then
                                q.Rows(CStr(va) & ":" & CStr(va)).Insert
                                q.Cells(va, 3) = "Wert"
                                q.Cells(va, tp) = "Auswahlwert"
                                q.Cells(va, 5) = p.Cells(n, m)
                                q.Cells(va, 4) = RP(q.Cells(va, 5))
                            End If
                            n = n + 1
                        Loop
                    End If
                Else
                    k = k + 1
                End If
            Else
                k = k + 1
            End If
        Loop
        ' If the boolean value b stayes "False" then it means that the attribute (ID) does not exist in the import file and has to be inserted.
        ' We include the attribute, attribute ID, its characteristics as well as default values if they exist.
        If b = False Then
            z = z + 1
            q.Cells(k, 1) = o.Cells(i, 12)
            q.Cells(k, 2) = "Attribut"
            q.Cells(k, 4) = o.Cells(i, 1)
            q.Cells(k, 5) = o.Cells(i, 11)
            If o.Cells(i, 10) = True Then q.Cells(k, di) = "x"
            If o.Cells(i, 8) <> "MerchandiseStyle" Then q.Cells(k, ar) = "x"
            q.Cells(k, tp) = o.Cells(i, 7)
            If o.Cells(i, 9) <> "" Then
                q.Cells(k, ui) = uns
                q.Cells(k, un) = o.Cells(i, 13)
            End If
            q.Cells(k, j) = "x"
            ' Insert default values
            If InStrRev(o.Cells(i, 7), "Wertemenge") > 0 Then
                l = FindColumn(p, o.Cells(i, 1), o.Cells(i, 11))
                k = k + 1
                m = 2
                Do Until p.Cells(m, l) = ""
                    q.Cells(k, 3) = "Wert"
                    s = p.Cells(m, l)
                    q.Cells(k, 5) = s
                    ' Create IDs for default values by using RP as RePlace functions
                    q.Cells(k, 4) = RP(s)
                    q.Cells(k, tp) = "Auswahlwert"
                    k = k + 1
                    m = m + 1
                Loop
            End If
        End If
        i = i + 1
    Loop
    
    ' Change format of all values to strings otherwise database has trouble to revognize default values correctly.
    q.Columns("D:E").NumberFormat = "@"
End Sub

Function RP(s)
    s = Replace(s, "Ä", "Ae")
    s = Replace(s, "Ö", "Oe")
    s = Replace(s, "Ü", "Ue")
    s = Replace(s, "ä", "ae")
    s = Replace(s, "ö", "oe")
    s = Replace(s, "ü", "ue")
    s = Replace(s, "ß", "ss")
    s = Replace(s, "é", "")
    s = Replace(s, " ", "")
    s = Replace(s, "-", "_")
    s = Replace(s, ",", "_")
    s = Replace(s, "+", "Plus")
    s = Replace(s, "/", "_")
    s = Replace(s, ".", "")
    s = Replace(s, "(", "")
    s = Replace(s, ")", "")
    s = Replace(s, "°", "Grad")
    s = Replace(s, "™", "")
    s = Replace(s, "®", "")
    s = Replace(s, "è", "")
    s = Replace(s, "%", "Prozent")
    
    
    RP = s
End Function

Sub FormCells(o)
    ' Goal: Change form of cells, necessary for database to recognize all content
    
    ' Variables
    Dim i As Integer
    Dim k As Long
    Dim ra(5) As Variant
    
    ra(0) = xlEdgeLeft
    ra(1) = xlEdgeTop
    ra(2) = xlEdgeRight
    ra(3) = xlEdgeBottom
    ra(4) = xlInsideHorizontal
    ra(5) = xlInsideHorizontal
    
    i = FindColumn(o, "Kommentar", "Kommentar")
    
    k = 3
    Do Until o.Cells(k, 1) = "" And o.Cells(k, 2) = "" And o.Cells(k, 3) = ""
        k = k + 1
    Loop
    k = k - 1
    
    ' What is needed to be done? Weight = xlHairline
    With o.Range(o.Cells(3, 1), o.Cells(k, i))
        For j = 0 To UBound(ra)
            With .Borders(ra(j))
                .LineStyle = xlContinuous
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .Weight = xlHairline
            End With
        Next
    End With
End Sub
