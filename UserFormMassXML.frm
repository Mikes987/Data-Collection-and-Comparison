VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormMassXML 
   Caption         =   "Mass Creation of XML Protocols"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8190
   OleObjectBlob   =   "UserFormMassXML.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormMassXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    ' Goal: Do mass data comparison
    
    ' Check if a directory has been chosen
    If DirectoryAddress.Caption = "" Or DirectoryAddress.Caption = "Falsch" Then
        MsgBox "Kein Verzeichnis gewählt"
        Exit Sub
    End If
    
    ' Apply-Button of single Data Comparison has been set to public to be recognized within this userform.
    
    ' Variables
    Dim s, s1, s2, t, u As String
    Dim path As String
    Dim i As Integer
    Dim a() As String
    
    ' Save path of directory and close userform
    path = DirectoryAddress.Caption
    Unload Me
    
    ' Load directories of attributes and dimensioned lists
    s1 = UserFormStart.PimPrimaryAddress.Caption
    s1 = Right(s1, Len(s1) - InStrRev(s1, "\"))
    s2 = UserFormStart.DIMAddress.Caption
    s2 = Right(s2, Len(s2) - InStrRev(s2, "\"))
    
    ' Create path with Dir() function and save all data paths into array
    t = Dir(path & "\*.xml")
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
    
    ' Now loop through this array and start data collection and comparison
    For i = 0 To UBound(a)
        s = path & "\" & a(i)
        UserFormStart.XMLAddress.Caption = s
        UserFormStart.ButtonApply_Click
        For Each Workbook In Workbooks
            ' In order to not have too many open files, close each protocol file
            Application.DisplayAlerts = False
            If Left(Workbook.Name, 4) = "PBK_" Then Workbook.Close
            Application.DisplayAlerts = True
        Next
    Next
    
    ' Close attribute and dimensioned list
    For Each Workbook In Workbooks
        If Workbook.Name = s1 Or Workbook.Name = s2 Then Workbook.Close
    Next
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonOpen_Click()
    ' Goal: Instead of choosing a single XML file, a directory will be chosen where all XML files are saved
    Dim Diafolder As FileDialog
    
    Set Diafolder = Application.FileDialog(msoFileDialogFolderPicker)
    Diafolder.AllowMultiSelect = False
    Diafolder.Show
    
    ' Load directory into caption
    DirectoryAddress.Caption = Diafolder.SelectedItems(1)
    'Set DiaFolder = Nothing
End Sub
