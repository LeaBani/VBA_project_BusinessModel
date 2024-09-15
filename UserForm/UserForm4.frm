VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "Empreinte carbone"
   ClientHeight    =   6420
   ClientLeft      =   30
   ClientTop       =   150
   ClientWidth     =   11355
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    Dim ctrl As Control
    
    Me.Width = 570
    Me.Height = 360
   
End Sub

Private Sub CommandButton1_Click()

    ValidateFormData UserForm4

End Sub

