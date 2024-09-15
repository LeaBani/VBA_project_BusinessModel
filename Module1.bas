Attribute VB_Name = "Module1"
Sub CollectTextBoxValues()

    Dim ctrl As Control
    Dim TextBoxValues As New Collection
    Dim i As Integer

    ' Parcourir tous les contrôles du UserForm
    For Each ctrl In Me.Controls
        ' Vérifier si le contrôle est un TextBox
        If TypeName(ctrl) = "TextBox" Then
            ' Ajouter la valeur du TextBox à la collection
            TextBoxValues.Add ctrl.Value, ctrl.Name
        End If
    Next ctrl

    ' Pour démontrer la collecte, afficher les valeurs dans une boîte de message
    For i = 1 To TextBoxValues.Count
        MsgBox "TextBox " & i & ": " & TextBoxValues(i)
    Next i
    
    Debug.Print (TextBoxValues)
    
End Sub

Sub DeleteRows()

    Dim ws As Worksheet
    Dim feuilleNames As Variant
    Dim feuilleName As Variant
    Dim lastRow As Long
    
    ' Liste des feuilles de calcul à traiter
    feuilleNames = Array("data_fluvial", "data_routier", "data_portuaire", "data_routier_preach")
    
    ' Traiter chaque feuille de calcul dans la liste
    For Each feuilleName In feuilleNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(feuilleName)
        If ws Is Nothing Then
            MsgBox "La feuille '" & feuilleName & "' n'existe pas.", vbCritical
        Else
            On Error GoTo 0
            ' Trouver la dernière ligne utilisée dans la feuille
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Supprimer les lignes de la dernière ligne jusqu'à la ligne 50 si nécessaire
            If lastRow >= 50 Then
                ws.Rows("50:" & lastRow).Delete
            Else
                ' Erreur, information utilisateur (désactivée)
                ' MsgBox "La feuille '" & feuilleName & "' ne contient pas assez de lignes pour supprimer depuis la ligne 50 jusqu'à la dernière ligne.", vbInformation
            End If
        End If
        On Error GoTo 0
    Next feuilleName

    ' Afficher un message de confirmation (désactivée)
    ' MsgBox "Les tableaux ont été initialisés."
End Sub

Sub ActualizeWithMaxValueToTransport()

    Dim ws As Worksheet
    Dim transportSheet As Worksheet
    Dim maxValue As Long
    Dim formulaRange As Range
    Dim feuilleNames As Variant
    Dim feuilleName As Variant
    
    Dim startTime As Double
    Dim endTime As Double
    Dim elapsedTime As Double
    
    ' Je supprime les lignes existantes
    Call DeleteRows
    
    ' Capturer le temps de début
    startTime = Timer

    ' Définir la feuille de calcul cout_transport
    On Error Resume Next
    Set transportSheet = ThisWorkbook.Sheets("cout_transport")
    If transportSheet Is Nothing Then
        MsgBox "La feuille 'cout_transport' n'existe pas.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Obtenir la valeur de la cellule D15 de la feuille cout_transport
    On Error Resume Next
    maxValue = transportSheet.Range("$D$18").Value
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'obtention de la valeur de la cellule D15 dans la feuille 'cout_transport'.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Vérifier si la valeur de E15 est valide
    If maxValue <= 0 Then
        MsgBox "La valeur de la cellule E18 dans la feuille 'cout_transport' est invalide.", vbCritical
        Exit Sub
    End If

    ' Liste des feuilles de calcul à traiter
    feuilleNames = Array("data_fluvial", "data_routier", "data_portuaire", "data_routier_preach")

    ' Traiter chaque feuille de calcul dans la liste
    For Each feuilleName In feuilleNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(feuilleName)
        If ws Is Nothing Then
            MsgBox "La feuille '" & feuilleName & "' n'existe pas.", vbCritical
        Else
            ' Définir la plage de la colonne D à remplir
            Set formulaRange = ws.Range("D2:D" & maxValue + 1)

            ' Tirer la formule de la cellule D2 vers le bas jusqu'à la dernière ligne spécifiée
            On Error Resume Next
            ws.Range("D2").AutoFill Destination:=formulaRange
            If Err.Number <> 0 Then
                MsgBox "Erreur lors du tirage de la formule de D2 dans la feuille '" & feuilleName & "'.", vbCritical
            End If
            On Error GoTo 0
        End If
        On Error GoTo 0
    Next feuilleName

        ' Capturer le temps de fin
    endTime = Timer
    
    ' Calculer le temps écoulé
    elapsedTime = endTime - startTime
    
    ' Afficher un message de confirmation avec le temps d'exécution
    MsgBox "Les tableaux ont été actualisés avec les données renseignées." & vbCrLf & _
           "Temps d'exécution : " & Format(elapsedTime, "0.00") & " secondes."
           
End Sub

Sub UserForm1_Show()

    UserForm1.Show
    ' Debug.Print "Show User form 1"
    
End Sub


Sub UserForm2_Show()

    UserForm2.Show
    
End Sub


Sub UserForm3_Show()

    UserForm3.Show
    
End Sub

Sub UserForm4_Show()

    UserForm4.Show
    
End Sub

' Visibilité des options lors de l'ouverture d'un UserForm

Sub UpdateVisibility(frameCtrl As MSForms.Frame, showTag As String, hideTag As String)

    Dim frameItem As Control
    
    ' Parcourir tous les contrôles dans la Frame spécifiée
    For Each frameItem In frameCtrl.Controls
        ' Afficher les contrôles avec le Tag correspondant
        If frameItem.Tag = showTag Then
            frameItem.Visible = True
        ' Cacher les contrôles avec le Tag correspondant
        ElseIf frameItem.Tag = hideTag Then
            frameItem.Visible = False
        End If
    Next frameItem
    
End Sub

Function ValidateAllTextBoxesInFrame(frm As MSForms.Frame) As Boolean

    Dim ctrl As Control
    
    ' Initialiser la fonction à True
    ValidateAllTextBoxesInFrame = True
    
    For Each ctrl In frm.Controls
        ' Vérifier si le contrôle est un TextBox et ne possède pas le tag "txt"
        If TypeOf ctrl Is MSForms.TextBox Then
            If ctrl.Tag <> "txt" Then
                If Not ValidateNumericTextBox(ctrl) Then
                    ValidateAllTextBoxesInFrame = False
                    Exit Function
                End If
            End If
        End If
    Next ctrl
    
End Function


' Procédure de validation des TextBox
Function ValidateNumericTextBox(txtBox As MSForms.TextBox) As Boolean

    ' Initialiser la fonction à True
    ValidateNumericTextBox = True
    
    If IsNumeric(txtBox.Value) Then
        txtBox.BackColor = RGB(255, 255, 255) ' Blanc
    Else
        txtBox.BackColor = RGB(247, 205, 201) ' Rouge clair
        MsgBox "Veuillez renseigner un nombre décimal dans le champ : " & txtBox.Name, vbExclamation
        ValidateNumericTextBox = False
    End If
    
End Function

Sub InitializeOptionsInUserForm(frm As UserForm)

    Dim ctrl As Control
    ' Parcourir tous les contrôles du UserForm
    For Each ctrl In frm.Controls
        ' Si le Tag du contrôle est "estimation", on le cache
        If ctrl.Tag = "estimation" Then
            ctrl.Visible = False
        End If
        If ctrl.Tag = "saisie" Then
            ctrl.Visible = False
        End If
    Next ctrl
    
End Sub



Sub ValidateFormData(frm As Object)

    Dim ctrl As Control
    Dim NomControle As String
    Dim ValeurControle As Variant
    Dim ws As Worksheet
    Dim NomRange As String
    
    ' Boucle à travers tous les contrôles du formulaire passé en paramètre (frm)
    For Each ctrl In frm.Controls
        ' Vérifier si le contrôle est un Frame
        If TypeName(ctrl) = "Frame" Then
            ' Appeler la validation des TextBox pour ce Frame
            If Not ValidateAllTextBoxesInFrame(ctrl) Then
                Exit Sub ' Sortir si la validation échoue
            End If
        End If
    Next ctrl

    ' Référence à la feuille "Input"
    Set ws = ThisWorkbook.Sheets("Input")

    ' Désactiver la mise à jour de l'écran et le calcul automatique pour accélérer l'exécution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Boucler à travers chaque contrôle dans le UserForm
    For Each ctrl In frm.Controls
        ' Récupérer le nom du contrôle
        NomControle = ctrl.Name
        
        ' Vérifier si le contrôle est de type TextBox, ComboBox, ListBox, OptionButton ou CheckBox
        Select Case TypeName(ctrl)
            Case "TextBox", "ComboBox", "ListBox", "OptionButton", "CheckBox"
                ' Construire le nom de la cellule cible
                NomRange = NomControle
                
                ' Récupérer la valeur du contrôle
                ValeurControle = ctrl.Value
                
                ' Vérifier si le NomRange existe dans la feuille "Input"
                On Error Resume Next
                If IsError(ws.Range(NomRange).Value) Then
                    ' Si le NomRange n'existe pas, afficher un message
                    MsgBox "La plage nommée '" & NomRange & "' n'a pas été trouvée dans la feuille 'Input'.", vbExclamation, "Plage non trouvée"
                Else
                    ' Si le NomRange existe, assigner la valeur à la cellule cible
                    If Not IsEmpty(ValeurControle) Then
                        If IsNumeric(ValeurControle) Then
                            ws.Range(NomRange).Value = CDbl(ValeurControle)
                        Else
                            ws.Range(NomRange).Value = ValeurControle
                        End If
                    End If
                End If
                On Error GoTo 0
            Case Else
                ' Ne rien faire pour les autres types de contrôles
        End Select
    Next ctrl

    ' Réactiver la mise à jour de l'écran et le calcul automatique
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Cacher le UserForm
    frm.Hide
    
End Sub
