VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Stockage Portuaire (optionnel)"
   ClientHeight    =   4690
   ClientLeft      =   -195
   ClientTop       =   -900
   ClientWidth     =   4770
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()

    ValidateFormData UserForm3

End Sub

Private Sub UserForm_Initialize()

    Dim ctrl As Control
    
    Me.Width = 370
    Me.Height = 400
   
End Sub

