Attribute VB_Name = "ModuleAttributeFiltern"
Sub AttributeFiltern(o)
    ' Goal: Filter Attribute list
    
    ' Variables
    Dim i, j, k, m As Integer
    Dim fi(7), s As String
    
    i = 1
    Do Until o.Cells(1, i) = ""
        i = i + 1
    Loop
    i = i - 1
    m = i
    
    ' what attributes do we want to keep? ==> End with: "_" & keywords:
    fi(0) = "Produkt"
    fi(1) = "Artikel"
    fi(2) = "DIM"
    fi(3) = "Steuerung"
    fi(4) = "Compliance"
    fi(5) = "Text"
    fi(6) = "produkt"
    fi(7) = "Dim"
    
    ' Remove attributes containing boolean values
    ' Certain attributes that have wrong descriptions and types
    ' attributes with wrong keywords
    
    i = 2
    Do Until o.Cells(i, 1) = ""
        s = o.Cells(i, 1)
        If _
            o.Cells(i, 4) = "Wahrheitswert" Or _
            (fi(0) <> Right(s, Len(s) - InStrRev(s, "_")) And _
            fi(1) <> Right(s, Len(s) - InStrRev(s, "_")) And _
            fi(2) <> Right(s, Len(s) - InStrRev(s, "_")) And _
            fi(3) <> Right(s, Len(s) - InStrRev(s, "_")) And _
            fi(4) <> Right(s, Len(s) - InStrRev(s, "_")) And _
            fi(5) <> Right(s, Len(s) - InStrRev(s, "_"))) Or _
            InStr(s, "_M_") > 0 Or _
            o.Cells(i, 1) = "Anlaesse_We_Steuerung" _
        Then
            o.Rows(CStr(i) & ":" & CStr(i)).Delete
        Else
            i = i + 1
        End If
    Loop
    
    o.Rows("1:1").AutoFilter
End Sub
