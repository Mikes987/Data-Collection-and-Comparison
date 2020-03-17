VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ButtonApply_Click()
    UserFormStart.Show
End Sub

Private Sub ButtonImport_Click()
    UserFormImport.Show
End Sub

Private Sub ButtonMassImport_Click()
    UserFormMassSingleImport.Show
End Sub

Private Sub ButtonMassXML_Click()
    UserFormMassXML.Show
End Sub

Private Sub ButtonNewPBK_Click()
    UserFormFilter.Show
End Sub


Private Sub ButtonPrimaryDoktrin_Click()
    UserFormPrimaryDoktrin.Show
End Sub


Private Sub CompPIMProtocoll_Click()
    UserFormCompPIMPro.Show
End Sub

Private Sub XMLLoad_Click()
    ' Goal: Before being able to download XML file, collect all product names from second sheet in column A and store in combobox. Addresses are in Column B
    
    Dim o As Object
    Dim i As Long
    Dim a() As String
    
    Set o = ThisWorkbook.Sheets("PBK-Doktrin-Addressen")
    
    i = 2
    Do Until o.Cells(i, 1) = ""
        If Not Not a Then
            ReDim Preserve a(i - 2)
            a(i - 2) = o.Cells(i, 1)
        Else
            ReDim a(i - 2)
            a(i - 2) = o.Cells(i, 1)
        End If
        i = i + 1
    Loop
    
    ' Store Array to combobox and show userform
    UserFormLoadXML.ComboBoxXML.List = a
    UserFormLoadXML.Show
End Sub
