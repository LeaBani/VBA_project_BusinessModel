Attribute VB_Name = "Module3"
Sub ResetData()
    ' D�clarez les variables
    Dim wsInput As Worksheet
    Dim initialValuesRange As Range
    Dim wsInputWasHidden As Boolean

    ' Affectez la feuille Input
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' V�rifiez si la feuille Input est masqu�e
    If wsInput.Visible = xlSheetHidden Or wsInput.Visible = xlSheetVeryHidden Then
        wsInputWasHidden = True
        wsInput.Visible = xlSheetVisible ' Rendre temporairement visible pour les modifications
    Else
        wsInputWasHidden = False
    End If

    ' Affectez la zone nomm�e "initial_values" dans la feuille Init
    On Error Resume Next ' G�rer l'erreur si la zone nomm�e n'existe pas
    Set initialValuesRange = ThisWorkbook.Names("initial_values").RefersToRange
    On Error GoTo 0 ' R�activez la gestion normale des erreurs

    If initialValuesRange Is Nothing Then
        MsgBox "La zone nomm�e 'initial_values' n'existe pas.", vbExclamation
        Exit Sub
    End If

    ' Effacez les donn�es existantes dans la feuille Input
    wsInput.Cells.ClearContents ' Efface uniquement le contenu, pas la mise en forme

    ' Copiez les valeurs de la zone nomm�e "initial_values" vers la feuille Input
    wsInput.Range("A1").Resize(initialValuesRange.Rows.Count, initialValuesRange.Columns.Count).Value = initialValuesRange.Value

    ' Si la feuille �tait masqu�e, la remasquer
    If wsInputWasHidden Then
        wsInput.Visible = xlSheetHidden
    End If

    MsgBox "Les donn�es ont �t� r�initialis�es avec succ�s depuis la zone 'initial_values'!"
End Sub

