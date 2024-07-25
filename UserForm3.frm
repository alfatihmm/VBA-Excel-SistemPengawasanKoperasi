VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "REKAPAN KEAKTIFAN KOPERASI"
   ClientHeight    =   8640.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10080
   OleObjectBlob   =   "UserForm3.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isPrintPreview As Boolean
Private Sub cbreset1_Click()
    Reset
End Sub

Private Sub CetakRekap_Click()
    Dim ws As Worksheet
    Set ws = keaktifan
    Dim iRow As Integer
    iRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    With ws.PageSetup
        .PrintArea = "A1:D99"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        
    End With
    Unload Me
    Unload MainMenu
    ws.PrintPreview
    MainMenu.Show
    UserForm3.Show
    
    
End Sub

Private Sub CommandButton1_Click()
    Dim ws As Worksheet
    Set ws = keaktifan
    Dim iRow As Integer
    iRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    With ws.PageSetup
        .PrintArea = "A1:D99"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        
    End With
    
    
     Unload Me
    Unload MainMenu
    ws.PrintOut
    MainMenu.Show
    UserForm3.Show
End Sub




Private Sub CommandButton2_Click()
Unload Me
MainMenu.Show
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub Naik_Click()
    SortDesc
End Sub

Private Sub Turun_Click()
SortByColumnI
End Sub




