VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ewsrekap 
   Caption         =   "REKAP EARLY WARNING SYSTEM"
   ClientHeight    =   8520.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12900
   OleObjectBlob   =   "ewsrekap.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ewsrekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btcetaksp_Click()
    Dim wsRekap As Worksheet
    Dim wsSP As Worksheet
    Dim fileName As String
    Dim savePath As Variant
    
    ' Mengatur worksheet yang sesuai
    Set wsRekap = ThisWorkbook.Sheets("rekapnilai")
    Set wsSP = ThisWorkbook.Sheets("sp")
    
    ' Menyimpan nilai ComboBox cbcetak ke sel N8 di sheet "rekapnilai"
    wsRekap.Range("N8").Value = Me.cbcetak.Value
    
    ' Menyimpan nilai ComboBox cbcetak ke sel O8 di sheet "sp"
    wsSP.Range("O8").Value = Me.cbcetak.Value
   

    With wsSP.PageSetup
        .PrintArea = "A1:J47"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
    End With
    
    NamaFile = TXTFOLDER.Value & "Surat Peringatan - " & wsSP.Range("O8").Value & ".Pdf"
    Sheet4.Select
    Sheet4.ExportAsFixedFormat Type:=xlTypePDF, _
    fileName:=NamaFile
    MsgBox "File telah disimpan dengan nama " & NamaFile
End Sub

Private Sub btprint_Click()
    Dim wsRekap As Worksheet
    Dim wsSP As Worksheet
       
    ' Mengatur worksheet yang sesuai
    Set wsRekap = ThisWorkbook.Sheets("rekapnilai")
    Set wsSP = ThisWorkbook.Sheets("sp")
    
    ' Menyimpan nilai ComboBox cbcetak ke sel N8 di sheet "rekapnilai"
    wsRekap.Range("N8").Value = Me.cbcetak.Value
    
    ' Menyimpan nilai ComboBox cbcetak ke sel O8 di sheet "sp"
    wsSP.Range("O8").Value = Me.cbcetak.Value
    

    ' Mengosongkan ComboBox
    Me.cbcetak.Value = ""
          
    With wsRekap.PageSetup
        .PrintArea = "A1:K81"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        
    End With
    Unload Me
    Unload MainMenu
    wsRekap.PrintPreview
    MainMenu.Show
    ewsrekap.Show
        
End Sub

Private Sub btrekap_Click()
    Dim wsRekap As Worksheet
    Dim wsSP As Worksheet
    Dim fileName As String
    Dim savePath As Variant
    
    ' Mengatur worksheet yang sesuai
    Set wsRekap = ThisWorkbook.Sheets("rekapnilai")
    Set wsSP = ThisWorkbook.Sheets("sp")
    
    ' Menyimpan nilai ComboBox cbcetak ke sel N8 di sheet "rekapnilai"
    wsRekap.Range("N8").Value = Me.cbcetak.Value
    
    ' Menyimpan nilai ComboBox cbcetak ke sel O8 di sheet "sp"
    wsSP.Range("O8").Value = Me.cbcetak.Value
   

    With wsRekap.PageSetup
        .PrintArea = "A1:K81"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
    End With
    
    NamaFile = TXTFOLDER.Value & "Hasil Pengawasan - " & wsSP.Range("O8").Value & ".Pdf"
    Sheet3.Select
    Sheet3.ExportAsFixedFormat Type:=xlTypePDF, _
    fileName:=NamaFile
    MsgBox "File telah disimpan dengan nama " & NamaFile
         
End Sub
Private Sub btfolder_Click()
    AturFolder
End Sub

Sub AturFolder()
    Dim SelectedFolder As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder"
        .ButtonName = "Confirm"
        If .Show = -1 Then
            SelectedFolder = .SelectedItems(1)
            MsgBox SelectedFolder
            TXTFOLDER.Value = SelectedFolder & "\"
        End If
    End With
End Sub

Private Sub btsp_Click()
Dim wsRekap As Worksheet
    Dim wsSP As Worksheet
       
    ' Mengatur worksheet yang sesuai
    Set wsRekap = ThisWorkbook.Sheets("rekapnilai")
    Set wsSP = ThisWorkbook.Sheets("sp")
    
    ' Menyimpan nilai ComboBox cbcetak ke sel N8 di sheet "rekapnilai"
    wsRekap.Range("N8").Value = Me.cbcetak.Value
    
    ' Menyimpan nilai ComboBox cbcetak ke sel O8 di sheet "sp"
    wsSP.Range("O8").Value = Me.cbcetak.Value
    

    ' Mengosongkan ComboBox
    Me.cbcetak.Value = ""
          
    With wsRekap.PageSetup
        .PrintArea = "A1:J42"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        
    End With
    Unload Me
    Unload MainMenu
    wsSP.PrintPreview
    MainMenu.Show
    ewsrekap.Show
End Sub




Private Sub CommandButton1_Click()
Unload Me
MainMenu.Show
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub lbews_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Dim daftar As Long
'Dim ws As Worksheet
'set ws as sheets("sheet1")

Set ws = Worksheets("nilai")


With Me.lbews
    If daftar < 2 Or daftar = ListCount Then
       daftar = .ListIndex
        cbcetak.Value = lbews.List(daftar, 1)
        
        'tbSTATUS.value = format(cdbl(tbSTATUS.value, "#;##0;00")
    End If
End With
End Sub


Private Sub previewews_Click()
    Dim ws As Worksheet
    Set ws = Sheet5
    Dim iRow As Integer
    iRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    With ws.PageSetup
        .PrintArea = "A1:D99"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .CenterHorizontally = True
        
    End With
    Unload Me
    Unload MainMenu
    ws.PrintPreview
    MainMenu.Show
    ewsrekap.Show
End Sub

Private Sub saveews_Click()
    Dim ws As Worksheet
    Set ws = Sheet5
    Dim iRow As Integer
    iRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    With ws.PageSetup
        .PrintArea = "A1:D99"
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .CenterHorizontally = True
        
    End With
    Unload Me
    Unload MainMenu
    ws.PrintOut
    MainMenu.Show
    ewsrekap.Show
End Sub


