Attribute VB_Name = "Module1"
Option Explicit
Sub urutkan()
    keaktifan.Range("A4:D12").Sort Key1:=Range("D4"), Order1:=xlAscending, Header:=xlYes
    
End Sub
Sub SortByColumnI()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Mengatur worksheet yang sesuai
    Set ws = ThisWorkbook.Sheets("nilai") ' Ganti dengan nama lembar kerja yang sesuai
    
    ' Menentukan baris terakhir dengan data di kolom I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Menentukan range yang akan diurutkan (A2:AA & lastRow)
    Dim sortRange As Range
    Set sortRange = ws.Range("A2:AA" & lastRow)
    
    ' Menentukan kolom yang digunakan sebagai kunci pengurutan (kolom I)
    Dim sortKey As Range
    Set sortKey = ws.Range("I2:I" & lastRow)
    
    ' Melakukan pengurutan berdasarkan kolom I secara menaik (Ascending)
    sortRange.Sort Key1:=sortKey, Order1:=xlAscending, Header:=xlNo
    
    
End Sub
Sub SortDesc()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Mengatur worksheet yang sesuai
    Set ws = ThisWorkbook.Sheets("nilai") ' Ganti dengan nama lembar kerja yang sesuai
    
    ' Menentukan baris terakhir dengan data di kolom I
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Menentukan range yang akan diurutkan (A2:AA & lastRow)
    Dim sortRange As Range
    Set sortRange = ws.Range("A2:AA" & lastRow)
    
    ' Menentukan kolom yang digunakan sebagai kunci pengurutan (kolom I)
    Dim sortKey As Range
    Set sortKey = ws.Range("I2:I" & lastRow)
    
    ' Melakukan pengurutan berdasarkan kolom I secara menaik (Ascending)
    sortRange.Sort Key1:=sortKey, Order1:=xlDescending, Header:=xlNo
    
    
End Sub

Sub Reset()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Mengatur worksheet yang sesuai
    Set ws = ThisWorkbook.Sheets("nilai") ' Ganti dengan nama lembar kerja yang sesuai
    
    ' Menentukan baris terakhir dengan data di kolom I
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Menentukan range yang akan diurutkan (A2:AA & lastRow)
    Dim sortRange As Range
    Set sortRange = ws.Range("A2:AA" & lastRow)
    
    ' Menentukan kolom yang digunakan sebagai kunci pengurutan (kolom I)
    Dim sortKey As Range
    Set sortKey = ws.Range("A2:A" & lastRow)
    
    ' Melakukan pengurutan berdasarkan kolom I secara menaik (Ascending)
    sortRange.Sort Key1:=sortKey, Order1:=xlAscending, Header:=xlNo
    
    
End Sub

Sub SImpan()
ThisWorkbook.Save
MsgBox "Data Berhasil Disimpan."
End Sub

