Attribute VB_Name = "ModuleSaveFiles"
Sub SaveFiles(wb, pbk)
    ' Goal: Create specific directories if non existent and save csv and xlsx files there
    
    Application.DisplayAlerts = False
    
    ' Variables
    Dim o As Object
    Dim fso, oFile As Object
    Dim i, j As Integer
    Dim s, t As String
    
    ' Create path for saving files and store into string variable
    s = ThisWorkbook.path & "\Ergebnisse" & "\" & Right(Date, Len(Date) - InStrRev(Date, ".")) & "_" & Mid(Date, InStr(Date, ".") + 1, 2) & "_" & Left(Date, 2)
    
    ' First save protocol
    t = "\Protokolle"
    t = s & t
    u = Dir(t, vbDirectory)
    If u = "" Then
        MkDir s
        MkDir t
    End If
    
    wb.SaveAs Filename:=t & "\" & "PBK_" & pbk & ".xlsx", FileFormat:=51
    
    ' Now create path for csv file
    t = "\CSV für Import"
    t = s & t
    u = Dir(t, vbDirectory)
    If u = "" Then
        MkDir t
    End If
    
    ' Create csv file
    Set o = wb.Sheets("PIM_Import")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.createtextfile(t & "\PBK_" & pbk & ".csv")
    
    ' Insert values into csv file
    i = 1
    Do Until o.Cells(i, 1) = ""
        j = 1
        Do Until o.Cells(i, j) = ""
            oFile.write o.Cells(i, j)
            ' In Germany, declaration of floating point numbers does not follow international standards.
            ' International: "1.0"
            ' Germany:       "1,0"
            ' Therefore, csv files often have another delimiter.
            ' Internatonal: ","
            ' Germany:      ";"
            If o.Cells(i, j + 1) <> "" Then oFile.write ";"
            j = j + 1
        Loop
        If o.Cells(i, j) = "" Then oFile.write vbNewLine
        i = i + 1
    Loop
    oFile.Close
    
    Application.DisplayAlerts = True
End Sub
