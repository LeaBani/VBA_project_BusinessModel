VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Solution mutlimodale route-fleuve"
   ClientHeight    =   9520.001
   ClientLeft      =   -3855
   ClientTop       =   -16320
   ClientWidth     =   7605
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

    ' Vérifiez l'état de la CheckBox (True ou False)
    If CheckBox1.Value = True Then
        ' Affiche Frame4
        Frame4.Visible = True
    Else
        ' Cache Frame4
        Frame4.Visible = False
        ResetTextBoxesInFrame Frame4
    End If
    
End Sub

Private Sub CheckBox2_Click()

    ' Vérifiez l'état de la CheckBox (True ou False)
    If CheckBox2.Value = True Then
        ' Affiche Frame6
        Frame6.Visible = True
    Else
        ' Cache Frame6
        Frame6.Visible = False
        ResetTextBoxesInFrame Frame6
    End If
    
End Sub

Private Sub CheckBox3_Click()

    ' Vérifiez l'état de la CheckBox (True ou False)
    If CheckBox3.Value = True Then
        ' Affiche Frame6
        Frame8.Visible = True
    Else
        ' Cache Frame6
        Frame8.Visible = False
        ResetTextBoxesInFrame Frame8
    End If
    
End Sub
Private Sub CheckBox4_Click()

    ' Vérifiez l'état de la CheckBox (True ou False)
    If CheckBox4.Value = True Then
        ' Affiche Frame6
        Frame9.Visible = True
    Else
        ' Cache Frame6
        Frame9.Visible = False
        ResetTextBoxesInFrame Frame9
    End If
    
End Sub

Private Sub ResetTextBoxesInFrame(frm As MSForms.Frame)

    Dim ctrl As Control
    ' Parcourt tous les contrôles dans le Frame spécifié
    For Each ctrl In frm.Controls
        
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Value = 0
        ElseIf TypeName(ctrl) = "ComboBox" Then
            ctrl.ListIndex = -1
        End If
        
    Next ctrl
    
End Sub

Private Sub CommandButton1_Click()
    
    
    ' Valider tous les TextBox dans les frames
    ' If Not ValidateAllTextBoxesInFrame(Frame4) Then Exit Sub
    ' If Not ValidateAllTextBoxesInFrame(Frame5) Then Exit Sub
    ' If Not ValidateAllTextBoxesInFrame(Frame6) Then Exit Sub
    ' If Not ValidateAllTextBoxesInFrame(unite_fluviale) Then Exit Sub
    
    
    ' Si toutes les validations réussissent, exécuter ValidateFormData
    ValidateFormData UserForm1
    
    ' Tester la cohérence des informations renseignées par l'utilisateur
    ' If Not ControlOfCoherenceA Then Exit Sub
    ' If Not ControlOfCoherenceB Then Exit Sub

End Sub

Private Function ControlOfCoherenceA() As Boolean

    '  Vérifier si ni "densite_aller" ni "nb_tonnes_A" n'est rempli
    If densite_aller.Value = 0 And nb_tonnes_A.Value = 0 Then
        ' Afficher un message d'erreur si les deux champs sont vides
        MsgBox "Veuillez choisir une option et renseigner une valeur pour la capacité d'emport de la barge pour chargeur A", vbExclamation, "Valeur requise"
        
        ' Arrêter l'exécution pour permettre à l'utilisateur de corriger
        ControlOfCoherenceA = False
        Exit Function
    End If
    ControlOfCoherenceA = True
    
End Function

Private Function ControlOfCoherenceB() As Boolean

    ' Vérifier si CheckBox2 est cochée
    If CheckBox2.Value = True Then

        '  PREMIER TEST Vérifier si ni "densite_retour" ni "nb_tonnes_B" n'est rempli
        If densite_retour.Value = 0 And nb_tonnes_B.Value = 0 Then
            ' Afficher un message d'erreur si les deux champs sont vides
            MsgBox "Veuillez choisir une option et renseigner une valeur pour la capacité d'emport de la barge pour chargeur B", vbExclamation, "Valeur requise"
            
            ' Arrêter l'exécution pour permettre à l'utilisateur de corriger
            ControlOfCoherenceB = False
            Exit Function
        End If
        
        ' DEUXIÈME TEST: Vérifier la CheckBox1 sur la Page 2
        MultiPage1.Value = 1 ' Rediriger vers Page2
        If CheckBox1.Value = False Then
            MsgBox "Veuillez renseigner les données du post-acheminement", vbExclamation, "Sélection requise"
            ControlOfCoherenceB = False
            Exit Function
        End If

    End If
    ControlOfCoherenceB = True

End Function


Private Sub OptionButton1_Click()

    UpdateVisibility Me.Frame5, "estimation", "saisie"
    
End Sub

Private Sub OptionButton2_Click()

    UpdateVisibility Me.Frame5, "saisie", "estimation"
    
End Sub

Private Sub OptionButton3_Click()

    UpdateVisibility Me.Frame6, "estimation", "saisie"

End Sub

Private Sub OptionButton4_Click()

    UpdateVisibility Me.Frame6, "saisie", "estimation"
    
End Sub


Private Sub OptionButton5_Click()

    UpdateVisibility Me.unite_fluviale, "saisie", "estimation"
    
End Sub

Private Sub OptionButton6_Click()

    UpdateVisibility Me.unite_fluviale, "estimation", "saisie"
    
End Sub

Private Sub OptionButton7_Click()

    UpdateVisibility Me.Frame4, "saisie", "estimation"

End Sub

Private Sub OptionButton8_Click()

    UpdateVisibility Me.Frame4, "estimation", "saisie"
    
End Sub

Private Sub OptionButton9_Click()

    UpdateVisibility Me.Frame3, "estimation", "saisie"

End Sub

Private Sub OptionButton10_Click()

    UpdateVisibility Me.Frame3, "saisie", "estimation"
    
End Sub



Private Sub UserForm_Initialize()

    Dim ctrl As Control
    
    ' Dimension du UserForm
    Me.Width = 570
    Me.Height = 500
    
    ' Initialisation des options
    InitializeOptionsInUserForm Me
    
    ' Cacher les parties optionnelles
    Frame6.Visible = False
    Frame4.Visible = False
    Frame8.Visible = False
    Frame9.Visible = False
    
    ResetTextBoxesInFrame Frame4
    ResetTextBoxesInFrame Frame6
    ResetTextBoxesInFrame Frame8
    ResetTextBoxesInFrame Frame9
       
   
End Sub

