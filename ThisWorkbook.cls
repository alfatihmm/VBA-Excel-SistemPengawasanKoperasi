VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Panggilmain()
    MainMenu.Show
End Sub
Sub PanggilUserForm2()
    UserForm2.Show
End Sub

Private Sub Workbook_Open()
    MainMenu.Show
End Sub
