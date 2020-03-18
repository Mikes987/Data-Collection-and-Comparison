VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormPrimaryDoktrin 
   Caption         =   "Compare protocol to primary Data Set"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8310.001
   OleObjectBlob   =   "UserFormPrimaryDoktrin.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormPrimaryDoktrin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    ' Goal: Compare data set of the created protocol by XML file with data set of primary and check mismatches.
    '       This is needed as the primary data set shall match our database but it is the XML file that gives us all inital information.
    
    ' Check if addresses are chosen
    If ProtocolAddress.Caption = "" Or ProtocolAddress.Caption = "Falsch" Then
        MsgBox "Kein Protokoll geladen"
        Exit Sub
    End If
    If PrimaryAddress.Caption = "" Or PrimaryAddress.Caption = "Falsch" Then
        MsgBox "Kein Primary-Lieferantenblatt angegeben"
        Exit Sub
    End If
    
    ' Variables
    Dim wb1, wb2 As Workbook
    Dim o, p As Object
    Dim s, t As String
    
    ' Save paths into string variables and close userform
    s = ProtocolAddress.Caption
    t = PrimaryAddress.Caption
    Unload Me
    
    ' Open and address workbooks and sheets
    Call LoadFile(wb1, o, ProtocolAddress.Caption, "Vergleich PIM - Doktrin")
    Call LoadFile(wb2, p, PrimaryAddress.Caption, 1)
    
    ' Do Check in subroutine
    Call CheckPrimary(o, p)
End Sub

Private Sub ButtonCancel_Click()
    Unload Me
End Sub

Private Sub ButtonLoadPrimarySheet_Click()
    PrimaryAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub

Private Sub ButtonLoadProtocol_Click()
    ProtocolAddress.Caption = Application.GetOpenFilename("Excel-Arbeitsmappe (*.xlsx), *.xlsx")
End Sub
