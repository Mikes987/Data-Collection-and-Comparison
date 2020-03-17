VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormFilter 
   Caption         =   "Attribute Filtern"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9090.001
   OleObjectBlob   =   "UserFormFilter.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButtonAppy_Click()
    ' Goal: Filter all attributes that do not follow the newly created ID standard
    
    ' Check if file has been chosen
    If PIMAddress.Caption = "" Or PIMAddress.Caption = "Falsch" Then
        MsgBox "Kein Blatt ausgewählt"
        Exit Sub
    End If
    
    ' Variables
    Dim wb As Workbook
    Dim o As Object
    Dim s, t As String
    Dim b, c As Boolean
    
    ' Open and address file
    Call LoadFile(wb, o, PIMAddress.Caption, 1)
    
    ' Close userform
    Unload Me
    
    ' Go through module
    Call AttributeFiltern(o)
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonChoose_Click()
    PIMAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub
