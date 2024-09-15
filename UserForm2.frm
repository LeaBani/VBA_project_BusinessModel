VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Solution 100 % route"
   ClientHeight    =   6360
   ClientLeft      =   -1710
   ClientTop       =   -6600
   ClientWidth     =   5925
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

    ' Vérifiez l'état de la CheckBox (True ou False)
    If CheckBox1.Value = True Then
        ' Affiche Frame4
        Frame2.Visible = True
    Else
        ' Cache Frame4
        Frame2.Visible = False
        
    End If
    
End Sub

Private Sub CommandButton1_Click()

    ValidateFormData UserForm2
    
End Sub


Private Sub routier_estim_A_Click()

    UpdateVisibility Me.Frame1, "estimation", "saisie"
    
End Sub

Private Sub routier_reel_A_Click()

    UpdateVisibility Me.Frame1, "saisie", "estimation"

End Sub

Private Sub routier_estim_B_Click()

    UpdateVisibility Me.Frame2, "estimation", "saisie"
    
End Sub

Private Sub routier_reel_B_Click()

    UpdateVisibility Me.Frame2, "saisie", "estimation"

End Sub

Private Sub UserForm_Initialize()

    Dim ctrl As Control

    Me.Width = 450
    Me.Height = 500
    
    ' Initialisation des options (Me fait référence à l'instance actuelle de l'objet)
    InitializeOptionsInUserForm Me
    
    ' Cacher les parties optionnelles
    Frame2.Visible = False
   
End Sub
