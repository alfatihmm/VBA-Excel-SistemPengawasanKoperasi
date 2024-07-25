VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "PROFIL KOPERASI"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13770
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cdCLEAR_Click()
    tbnama.Value = ""
    tbnomor.Value = ""
    tbtanggal.Value = ""
    tbalamat.Value = ""
    tbtelp.Value = ""
    tbbentuk.Value = ""
    tbjenis.Value = ""
    tbnib.Value = ""
    tbnik.Value = ""
    tbnpwp.Value = ""
End Sub

Private Sub cdEDIT_Click()
Set ws = Worksheets("datakop")
Dim rng As Range

If MsgBox("Apakah Data Akan Diperbarui ?", vbOKCancel + vbQuestion, "konfirmasi") = vbOK Then
    Set rng = ws.Range("B2:B200").Find(tbnama.Value, LookIn:=xlValues)
 

        If Not rng Is Nothing Then
            Baris = rng.Row
          
With ws
.Cells(Baris, 2).Value = tbnama.Value
.Cells(Baris, 3).Value = tbnomor.Value
.Cells(Baris, 4).Value = tbtanggal.Value
.Cells(Baris, 5).Value = tbalamat.Value
.Cells(Baris, 6).Value = tbtelp.Value
.Cells(Baris, 7).Value = tbbentuk.Value
.Cells(Baris, 8).Value = tbjenis.Value
.Cells(Baris, 9).Value = tbnib.Value
.Cells(Baris, 10).Value = tbnik.Value
.Cells(Baris, 11).Value = tbnpwp.Value


End With

End If
End If
With UserForm1
        .tbnama.Value = ""
        .tbnomor.Value = ""
        .tbtanggal.Value = ""
        .tbalamat.Value = ""
        .tbtelp.Value = ""
        .tbbentuk.Value = ""
        .tbjenis.Value = ""
        .tbnib.Value = ""
        .tbnik.Value = ""
        .tbnpwp.Value = ""
    End With
End Sub

Private Sub cdHAPUS_Click()
    ' fungsi hapus disini = kalau kita mendelete baris didata excel
    ' jadi blok area database buat sebanyak mungkin kebawah
    ' jika blok area sampai baris 10, kemudian baris 10 dihapus, maka input baru
    ' akan masuk ke baris 10, namun dilist box (dengan rowssource "hasil") tidak akan nampak
    
    nama = tbnama.Value
    ' ini perlu dibuat sebagai acuan pertama dan dipakai pada script berikutnya
    ' jika masterdata berupa nama stok, maka nama ini akan merujuk pada TextBoxStok ("nama = tbstock")
    ' atau barang = tbstok
    
    Set ws = Worksheets("datakop")
    
    With ws.Range("B2:B200")
        Set c = .Find(nama, LookIn:=xlValues)
        Baris = c.Row
    
        If MsgBox("Apakah Data Akan Dihapus?", vbOKCancel + vbQuestion, "Konfirmasi") = vbOK Then
            If Not c Is Nothing Then
                ws.Cells(Baris, 2).EntireRow.Delete
            End If
            MsgBox "Data Telah Dihapus"
        End If
    End With
End Sub


Private Sub cdSIMPAN_Click()

Set ws = Worksheets("datakop")


With ws.Range("B2:K2")

barisakhir = ws.Range("B" & Rows.Count).End(xlUp).Row + 1

ws.Range("B" & barisakhir).Value = tbnama.Text
ws.Range("C" & barisakhir).Value = tbnomor.Text
ws.Range("D" & barisakhir).Value = tbtanggal.Text
ws.Range("E" & barisakhir).Value = tbalamat.Text
ws.Range("F" & barisakhir).Value = tbtelp.Text
ws.Range("G" & barisakhir).Value = tbbentuk.Text
ws.Range("H" & barisakhir).Value = tbjenis.Text
ws.Range("I" & barisakhir).Value = tbnib.Text
ws.Range("J" & barisakhir).Value = tbnik.Text
ws.Range("K" & barisakhir).Value = tbnpwp.Text

End With

 ' Menampilkan notifikasi data telah disimpan
    MsgBox "Data telah disimpan."
    
    ' Mengosongkan TextBox setelah tombol ditekan
    With UserForm1
        .tbnama.Value = ""
        .tbnomor.Value = ""
        .tbtanggal.Value = ""
        .tbalamat.Value = ""
        .tbtelp.Value = ""
        .tbbentuk.Value = ""
        .tbjenis.Value = ""
        .tbnib.Value = ""
        .tbnik.Value = ""
        .tbnpwp.Value = ""
    End With
End Sub


Private Sub lbHasilInput_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Dim daftar As Long
'Dim ws As Worksheet
'set ws as sheets("sheet1")

Set ws = Worksheets("datakop")


With Me.lbHasilInput
    If daftar < 2 Or daftar = ListCount Then
       daftar = .ListIndex
        tbnama.Value = lbHasilInput.List(daftar, 1)
        tbnomor.Value = lbHasilInput.List(daftar, 2)
        tbtanggal.Value = lbHasilInput.List(daftar, 3)
        tbalamat.Value = lbHasilInput.List(daftar, 4)
        tbtelp.Value = lbHasilInput.List(daftar, 5)
        tbbentuk.Value = lbHasilInput.List(daftar, 6)
        tbjenis.Value = lbHasilInput.List(daftar, 7)
        tbnib.Value = lbHasilInput.List(daftar, 8)
        tbnik.Value = lbHasilInput.List(daftar, 9)
        tbnpwp.Value = lbHasilInput.List(daftar, 10)

        'tbSTATUS.value = format(cdbl(tbSTATUS.value, "#;##0;00")
    End If
End With
End Sub

Private Sub UserForm_Terminate()
    MainMenu.Show
End Sub
