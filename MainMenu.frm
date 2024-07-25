VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   9975.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17520
   OleObjectBlob   =   "MainMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btexit_Click()
 Unload Me
 Application.Quit
 Application.DisplayAlerts = False
 ThisWorkbook.Saved = True
 
End Sub

Private Sub CommandButton1_Click()
  UserForm1.Show
  
End Sub

Private Sub CommandButton2_Click()
    UserForm2.Show
End Sub

Private Sub CommandButton3_Click()
    UserForm3.Show
End Sub

Private Sub CommandButton4_Click()
    ewsrekap.Show
End Sub

Private Sub CommandButton5_Click()
    SImpan
End Sub

Private Sub UserForm_Activate()
ActiveWindow.WindowState = xlMaximized
With Me
   .Height = Application.Height
   .Width = Application.Width
   .Left = -6
   .Top = 0
End With
End Sub

Private Sub UserForm_Click()

End Sub
