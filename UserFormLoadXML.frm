VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormLoadXML 
   Caption         =   "XML Doktrin Laden"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5625
   OleObjectBlob   =   "UserFormLoadXML.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormLoadXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonLoad_Click()
    ' Goal: Load XML file for product chosen from combobox
        
    ' Variables
    Dim o As Object
    Dim i As Integer
    Dim s, t As String
    
    ' Address information of userform
    Set o = ThisWorkbook.Sheets("PBK-Doktrin-Addressen")
    s = UserFormLoadXML.ComboBoxXML.Text
    
    ' Check if a product has been chosen
    If s = "" Then
        MsgBox "No product chosen"
        Exit Sub
    End If
    
    ' Quit userform
    Unload Me
    
    ' Look for position of product name
    i = 2
    Do Until o.Cells(i, 1) = s
        i = i + 1
    Loop
    
    ' When matched, save address stored in column B
    t = o.Cells(i, 2)
    
    ' We need further objects to prepare download
    Dim p As Object
    Dim u As String
    
    Set p = CreateObject("MSXML2.DOMDocument")
    
    ' Set path for download
    u = "C:\User\data\..." ' Only example path
    u = u & s & ".xml"
    
    ' Do Download
    With p
        .async = False
        .validateonparse = False
        .Load t
        .Save u
    End With
End Sub
