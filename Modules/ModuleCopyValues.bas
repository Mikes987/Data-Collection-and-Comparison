Attribute VB_Name = "ModuleCopyValues"
Sub CopyValues(wb, o, r, at, va)
    ' Goal: Go through parsed XML file again and copy default values
    
    ' o:   XML file
    ' r:   protocol
    
    ' Variables
    Dim i, j, k, m, n As Integer
    Dim s As String
    
    ' Another sheet will be created in the new file in which default values will be inserted.
    ' Row 1 contains attribute names, default values will begin in row 2.
    wb.Sheets.Add after:=r
    Set q = wb.ActiveSheet
    q.Name = "Vorgabewerte"
    
    ' For Performance use Screen-Updating-Function as False
    Application.ScreenUpdating = False
    
    i = 2
    j = 1
    Do Until r.Cells(i, 2) = ""
        ' Datatype tells us if there will be default values. If yes, then copy the attribute ID into row 1 of the new sheet.
        s = r.Cells(i, 7)
        If InStrRev(s, "Wertemenge") > 0 Then
            q.Cells(1, j) = r.Cells(i, 1)
            ' Now we gow through the XML file again. We need 2 columns, attribute name and values
            ' at: Column with attribute name
            ' va: Column with default values
            k = 2
            m = 2
            Do Until o.Cells(m, 1) = ""
                If o.Cells(m, at) = r.Cells(i, 11) And o.Cells(m, va) <> "" Then
                    q.Cells(k, j) = o.Cells(m, va)
                    k = k + 1
                End If
                m = m + 1
            Loop
            j = j + 1
        End If
        i = i + 1
    Loop
    q.Rows("1:1").Font.Bold = True
    q.Cells.EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
End Sub
