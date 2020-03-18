VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMassSingleImport 
   Caption         =   "Create Mass Import"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7290
   OleObjectBlob   =   "UserFormMassSingleImport.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormMassSingleImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    ' Goal: Instead of creating import of single protocol file, load a directory and loop through all protocol files
    '       That is why all protocols are saved automatically in one directory, depending on date when they have been created.
    '       Simply load that directory.
    
    ' Check if Directory has been chosen
    If DirectoryAddress.Caption = "" Or DirectoryAddress.Caption = "Falsch" Then
        MsgBox "Kein Verzeichnis geladen"
        Exit Sub
    End If
    
    ' Variables
    Dim wb1, wb2, wb3 As Workbook
    Dim o, p, q As Object
    Dim s, t, u As String
    Dim a() As String
    Dim i As Integer
    
    
    ' Save path of directory and close userform
    s = DirectoryAddress.Caption
    Unload Me
    
    
    ' Load directory and loop through all files, save name of files into array
    t = Dir(s & "\*.xlsx")
    ReDim a(0)
    a(0) = t
    t = Dir
    i = 1
    Do While t <> ""
        ReDim Preserve a(i)
        a(i) = t
        i = i + 1
        t = Dir
    Loop
    
    ' Create Import file. This file will actually ignore if an import file is already open Since I don't know how big this file can be then
    Workbooks.Add
    Set wb3 = ActiveWorkbook
    Set q = wb3.ActiveSheet
    Call PrepareImport(q)
    
    ' We load the file, 2 sheets in that file, protocol and default values.
    ' However, we do not go specifically through single import algorithm but simply call subroutine Import
    For i = 0 To UBound(a)
        u = s & "\" & a(i)
        Call LoadFile(wb1, o, u, "Vergleich PIM - Doktrin")
        Call LoadFile(wb2, p, u, "Vorgabewerte")
        Call Import(o, p, q)
    Next
    
    ' As a last step, form cells
    Call FormCells(q)
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonLoad_Click()
    ' Goal: Instead of loading single file, load entire directory.
    
    ' Variables
    Dim Diafolder As FileDialog
    
    Set Diafolder = Application.FileDialog(msoFileDialogFolderPicker)
    Diafolder.AllowMultiSelect = False
    Diafolder.Show
    
    DirectoryAddress.Caption = Diafolder.SelectedItems(1)
    Set Diafolder = Nothing
End Sub
