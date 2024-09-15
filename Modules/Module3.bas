Attribute VB_Name = "Module3"
Sub ResetData()
    ' Déclarez les variables
    Dim wsInput As Worksheet
    Dim initialValuesRange As Range
    Dim wsInputWasHidden As Boolean

    ' Affectez la feuille Input
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Vérifiez si la feuille Input est masquée
    If wsInput.Visible = xlSheetHidden Or wsInput.Visible = xlSheetVeryHidden Then
        wsInputWasHidden = True
        wsInput.Visible = xlSheetVisible ' Rendre temporairement visible pour les modifications
    Else
        wsInputWasHidden = False
    End If

    ' Affectez la zone nommée "initial_values" dans la feuille Init
    On Error Resume Next ' Gérer l'erreur si la zone nommée n'existe pas
    Set initialValuesRange = ThisWorkbook.Names("initial_values").RefersToRange
    On Error GoTo 0 ' Réactivez la gestion normale des erreurs

    If initialValuesRange Is Nothing Then
        MsgBox "La zone nommée 'initial_values' n'existe pas.", vbExclamation
        Exit Sub
    End If

    ' Effacez les données existantes dans la feuille Input
    wsInput.Cells.ClearContents ' Efface uniquement le contenu, pas la mise en forme

    ' Copiez les valeurs de la zone nommée "initial_values" vers la feuille Input
    wsInput.Range("A1").Resize(initialValuesRange.Rows.Count, initialValuesRange.Columns.Count).Value = initialValuesRange.Value

    ' Si la feuille était masquée, la remasquer
    If wsInputWasHidden Then
        wsInput.Visible = xlSheetHidden
    End If

    MsgBox "Les données ont été réinitialisées avec succès depuis la zone 'initial_values'!"
End Sub

