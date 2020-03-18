VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormImport 
   Caption         =   "Prepare Import for Product"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7785
   OleObjectBlob   =   "UserFormImport.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    ' Goal: Create template for import of products by placing attributes and default values from protocols into the correct position.
    '       Multiple protocols can be loaded into one import file.
    
    ' Variables
    Dim wb1, wb2, wb3, wb4 As Workbook
    Dim o, p, q As Object
    Dim s As String
    Dim n1, n2, n3 As String
    Dim b As Boolean
    
    ' Sheet names of protocol - hopefully nobody changes them manually
    n1 = "Vergleich PIM - Doktrin"
    n2 = "Vorgabewerte"
    n3 = "PBK-Import_PIM"
    
    ' Open and address protocol sheets
    Call Check(wb1, o, ProtocolAddress.Caption, n1)
    Call Check(wb2, p, ProtocolAddress.Caption, n2)
    
    ' Alle Informationen vorhanden, Benutzeroberfläche kann geschlossen werden.
    Unload Me
    
    ' Since any import file for database can contain a high number of products, the user has two options:
    ' 1. No import file is open, a new file will be generated and automatically saved at the end. Its sheet name will be n3
    ' 2. An import file has already been genereated and is still open. In that case, the script will find that file and add further product information
    ' into that file.
    b = False
    For Each Workbook In Workbooks
        For Each Worksheet In Workbook.Worksheets
            If Worksheet.Name = n3 Then
                b = True
                Set wb3 = Workbook
                Set q = wb3.Sheets(n3)
                If q.Cells(1, 1) = "" Then Call PrepareImport(q)
                Exit For
            End If
        Next
        If b = True Then Exit For
    Next
    If b = False Then
        Workbooks.Add
        Set wb3 = ActiveWorkbook
        Set q = wb3.ActiveSheet
        q.Name = n3
        Call PrepareImport(q)
    End If
    
    ' Do import here
    Call Import(o, p, q)
    Call FormCells(q)
    
    ' Save
    Application.DisplayAlerts = False
    wb3.SaveAs Filename:=ThisWorkbook.path & "\Import_PIM.xlsx"
    Application.DisplayAlerts = True
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonImport_Click()
    ImportAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*xlsx), *.xlsx")
End Sub

Private Sub ButtonProtocol_Click()
    ProtocolAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub



Sub Check(wb, o, s, n)
    ' Open and address files, general function
    Dim t As String
    Dim b As Boolean
    
    t = Right(s, Len(s) - InStrRev(s, "\"))
    b = False
    For Each Workbook In Workbooks
        If Workbook.Name = t Then
            b = True
            Set wb = Workbook
            Set o = wb.Sheets(n)
            Exit For
        End If
    Next
    If b = False Then
        Workbooks.Open s
        Set wb = ActiveWorkbook
        Set o = wb.Sheets(n)
    End If
End Sub
