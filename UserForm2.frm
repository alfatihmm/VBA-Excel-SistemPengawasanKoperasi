VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "INPUT NILAI KOPERASI"
   ClientHeight    =   9600.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13560
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbsbmit_Click()
    Dim ws As Worksheet
    Dim i As Integer
    Dim lastRow As Long
    
    ' Mengatur worksheet yang sesuai
    Set ws = ThisWorkbook.Sheets("nilai")
    
    ' Menentukan baris terakhir dengan data
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    
    ' Menambahkan nomor urut pada kolom A
    ws.Cells(lastRow, "A").Value = lastRow - 1
    
    ' Menyimpan nilai ComboBox cbnama ke kolom B
    ws.Cells(lastRow, "B").Value = Me.cbnama.Value
    
    ' Menyimpan data ComboBox tk1 hingga tk10 ke baris selanjutnya, mulai dari kolom C
    For i = 1 To 10
        If i = 7 Then
            ' Tetap sesuai dengan input ComboBox tk7
            ws.Cells(lastRow, i + 2).Value = Me.Controls("tk7").Value
        Else
            If Me.Controls("tk" & i).Value = "Memenuhi" Then
                ws.Cells(lastRow, i + 2).Value = 1
            Else
                ws.Cells(lastRow, i + 2).Value = 0
            End If
        End If
    Next i
    
    ' Menyimpan nilai dari TextBox pr1 hingga pr2, kk1 hingga kk10, dan modal1 hingga modal3
    ws.Cells(lastRow, 13).Value = Me.pr1.Value
    ws.Cells(lastRow, 14).Value = Me.pr2.Value
    
    For i = 1 To 10
        ws.Cells(lastRow, 14 + i).Value = Me.Controls("kk" & i).Value
    Next i
    
    ws.Cells(lastRow, 25).Value = Me.modal1.Value
    ws.Cells(lastRow, 26).Value = Me.modal2.Value
    ws.Cells(lastRow, 27).Value = Me.modal3.Value
    
    ' Menampilkan notifikasi
    MsgBox "Data telah disimpan."
    
    ' Mengosongkan ComboBox dan TextBox
    Me.cbnama.Value = ""
    Me.pr1.Value = ""
    Me.pr2.Value = ""
    
    For i = 1 To 10
        Me.Controls("tk" & i).Value = ""
        Me.Controls("kk" & i).Value = ""
    Next i

    Me.modal1.Value = ""
    Me.modal2.Value = ""
    Me.modal3.Value = ""
End Sub
Private Sub UserForm2_Terminate()
    MainMenu.Show
End Sub

Private Sub UserForm_Click()

End Sub
