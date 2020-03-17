VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormStart 
   Caption         =   "Data Comparison"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10365
   OleObjectBlob   =   "UserFormStart.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub ButtonApply_Click()
    ' Goal: Create a List with all attributes and their description for their product
    
    ' Check files are loaded
    If XMLAddress.Caption = "" Or XMLAddress.Caption = "Falsch" Then
        MsgBox "XML-Datei nicht angegeben"
        Exit Sub
    End If
    If PimPrimaryAddress.Caption = "" Or PimPrimaryAddress.Caption = "Falsch" Then
        MsgBox "PIM-Attribute_Primary Datei nicht angewählt"
        Exit Sub
    End If
    If DIMAddress.Caption = "" Or DIMAddress.Caption = "Falsch" Then
        MsgBox "Blatt mit dimensionierten Werten nicht angewählt"
        Exit Sub
    End If
    
    ' Variables
    Dim wb1, wb2, wb3, wb4 As Workbook
    Dim n, o, p, q, r As Object
    Dim pbk As String
    Dim s, t, u As String
    Dim b, b1 As Boolean
    Dim a() As String
    Dim num(9) As String
    Dim i, k As Integer
    
    ' o:   XML File parsed into Excel file
    ' p:   Attribute-List
    ' pbk: pbk-Label or product name
    
    ' Open and address files
    Call LoadXML(wb1, o, XMLAddress.Caption)
    Call LoadFile(wb2, p, PimPrimaryAddress.Caption, 1)
    Call LoadFile(wb3, q, DIMAddress.Caption, 1)
    
    ' Close userform
    Unload Me
    
    ' We have to go over the parsed XML file twice so we reference the columns here
    Dim pb As Integer ' PBK-label or name
    Dim at As Integer ' Attribut name
    Dim mv As Integer ' contains defaultvalues and if yes, single or multi?
    Dim pf As Integer ' Mandatory?
    Dim tp As Integer ' Datatype, String, Integer, ...
    Dim va As Integer ' Default values if they exist
    Dim td As Integer ' Item, ItemOption or MerchandiseStyle
    Dim gr As Integer ' Product group
    Dim st As Integer ' if 0 ==> Compliance
    Dim uo As Integer ' Unit, full (kilogramm)
    Dim un As Integer ' Unit, physical (kg)
    
    pb = FindColumn(o, "code", "PBK-Name")
    mv = FindColumn(o, "multivalued", "Wertemenge, mehrfach")
    pf = FindColumn(o, "mandatory", "Pflichteintrag")
    st = FindColumn(o, "sortKey4", "Steuerung")
    at = FindColumn(o, "ns1:Text6", "Attribut")
    td = FindColumn(o, "name", "Artikel-/Produktebene")
    tp = FindColumn(o, "ns1:DataType", "Datentyp")
    gr = FindColumn(o, "ns1:GroupCode", "Gruppenbezeichnung")
    va = FindColumn(o, "ns1:ValidValue", "Vorgabewerte")
    un = FindColumn(o, "symbol", "Einheit, physikalisch")
    uo = un - 1
    
    ' Sometimes the pbk label or name contains prefixes that have to be removed. We know that these prefixes only contain numbers
    num(0) = "0"
    num(1) = "1"
    num(2) = "2"
    num(3) = "3"
    num(4) = "4"
    num(5) = "5"
    num(6) = "6"
    num(7) = "7"
    num(8) = "8"
    num(9) = "9"
    
    ' Check for prefixes
    pbk = o.Cells(2, pb)
    If InStr(pbk, "_") > 0 Then
        u = pbk
        b1 = False
        Do While b = False And InStrRev(u, "_") > 0
            k = InStrRev(u, "_") - 1
            t = Mid(u, k, 1)
            For i = 0 To UBound(num)
                If num(i) = t Then
                    b = True
                    Exit For
                End If
            Next
            If b = False Then
                u = Left(u, k)
            Else
                pbk = Right(pbk, Len(pbk) - k - 1)
            End If
        Loop
    End If
    
    ' Next we want to create a new list. This is important as this list is the first step for import into database.
    ' XML information will first be stored into an array through the function DoSearch()
    a = DoSearch(o, p, at, mv, pf, tp, va, td, gr, st, uo, un)
    
    ' Create a new file where the content of the array will be stored
    Workbooks.Add
    Set wb4 = ActiveWorkbook
    Set r = wb4.ActiveSheet
    r.Cells(1, 1).Activate
    ActiveWindow.WindowState = xlMaximized
    r.Name = "Vergleich PIM - Doktrin"
    
    ' Module List:       Insert content of array into new file
    ' Module ProdArt:    Check if product- or article based product
    ' Module Check:      Do Comparison with attribute and DIM list
    ' Module CopyValues: Copy default values

    Call List(a, pbk, r)
    Call prodart(q, r)
    Call Check(wb4, p, r)
    Call CopyValues(wb4, o, r, at, va)
    
    ' Done, XML file can be closed
    Application.DisplayAlerts = False
    wb1.Close
    Application.DisplayAlerts = True
    
    ' Save as xlsx and csv file
    Call SaveFiles(wb4, pbk)
    
    r.Range("A1:P1").AutoFilter
    r.Activate
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonDim_Click()
    DIMAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub


Private Sub ButtonReadList_Click()
    PBKNoAttAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub


Private Sub ButtonReadPIM_Click()
    PimPrimaryAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub


Private Sub ButtonReadXML_Click()
    XMLAddress.Caption = Application.GetOpenFilename("XML-Datei (*.xml), *.xml")
End Sub
