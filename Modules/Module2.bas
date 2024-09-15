Attribute VB_Name = "Module2"
Sub RefreshPowerPivotCharts()

    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim pvtTable As PivotTable
    
    ' Rafraîchir le modèle PowerPivot
    ThisWorkbook.Model.Refresh
    
    ' Boucler sur toutes les feuilles
    For Each ws In ThisWorkbook.Worksheets
        ' Boucler sur tous les TCD de la feuille
        For Each pvtTable In ws.PivotTables
            ' Rafraîchir le TCD
            pvtTable.RefreshTable
        Next pvtTable
        
        ' Boucler sur tous les graphiques de chaque feuille
        For Each chtObj In ws.ChartObjects
            ' Forcer Excel à redessiner le graphique en activant puis désactivant la feuille
            ws.Activate
            chtObj.Chart.Refresh
        Next chtObj
    Next ws
    
    CheckIfMeasureInValues
    
    MsgBox "Les graphiques et TCD liés au modèle PowerPivot ont été mis à jour avec succès!", vbInformation
    
End Sub

Sub PrintDashboardAsPdf()

    Dim wsSource As Worksheet
    Dim dashboardRange As Range
    Dim pdfFilePath As String
    Dim currentWorkbookPath As String
    
    Dim wsStart As Worksheet
    
    Set wsStart = ThisWorkbook.Sheets("Start")
    
    ' Rester sur l'onglet Start
    wsStart.Activate
    
    RefreshPowerPivotCharts
    
    CheckIfReturn
        
    FiltrerTCD
    
    ' Désactiver les fonctionnalités pour améliorer les performances
    Application.ScreenUpdating = False ' Désactiver la mise à jour de l'écran
    Application.Calculation = xlCalculationManual ' Désactiver le calcul automatique
    Application.DisplayAlerts = False ' Désactiver les alertes
    Application.EnableEvents = False ' Désactiver les événements
    

    ' Définir la feuille source
    Set wsSource = ThisWorkbook.Sheets("Analyse")
    
    ' Temporiser la feuille Analyse pour l'exportation
    wsSource.Visible = xlSheetVisible ' Rendre l'onglet visible temporairement
    
    ' Définir la plage nommée
    On Error Resume Next
    Set dashboardRange = wsSource.Range("Dashboard")
    On Error GoTo 0
    
    ' Vérifier si la plage nommée existe
    If dashboardRange Is Nothing Then
        MsgBox "La plage nommée 'Dashboard' n'existe pas sur la feuille 'Analyse'.", vbExclamation
        Exit Sub
    End If
    
    ' Obtenir le chemin du dossier du classeur actif
    currentWorkbookPath = ThisWorkbook.Path
    If currentWorkbookPath = "" Then
        MsgBox "Le classeur doit être enregistré avant de procéder.", vbExclamation
        Exit Sub
    End If
    
    ' Définir le chemin du fichier PDF
    pdfFilePath = currentWorkbookPath & "\Dashboard.pdf"
    
    ' Définir la zone d'impression
    wsSource.PageSetup.PrintArea = dashboardRange.Address
    
    ' Configurer la mise en page en paysage
    With wsSource.PageSetup
        .Orientation = xlLandscape
        .Zoom = False ' Utiliser les dimensions spécifiées pour l'impression
        .FitToPagesWide = 1 ' Ajuster automatiquement la largeur pour ne pas couper en mettant sur False
        .FitToPagesTall = 1
    End With
    
    ' Exporter la plage en PDF
    wsSource.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard
    
    ' Ouvrir le fichier PDF
    Shell "explorer.exe " & pdfFilePath, vbNormalFocus
    
    ' Réactiver l'onglet Start
    wsStart.Activate
    
    ' Assurez-vous que l'onglet Analyse reste masqué
    wsSource.Visible = xlSheetHidden  ' Garder l'onglet Analyse masqué
    
    ' Réactiver les fonctionnalités désactivées
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' Avertir l'utilisateur
    MsgBox "La plage 'Dashboard' a été enregistrée en tant que PDF et le fichier a été ouvert.", vbInformation
    
End Sub

Sub FiltrerTCD()

    Dim ws As Worksheet
    
    ' Référence à la feuille contenant le TCD
    Set ws = ThisWorkbook.Sheets("Analyse")
    
    ' Lister tous les graphiques
    ' For Each obj In ws.ChartObjects
        ' Debug.Print "Graphique: " & obj.Name
    ' Next obj
    
    ' Vérifier si la cellule F6 est égale à 0
    If ws.Range("D7").Value = 0 Then
        ' Masquer le Graphique7
        ws.ChartObjects("Chart 7").Visible = False
    Else
        ' Afficher le Graphique7 et exécuter FiltrerTCD
        ws.ChartObjects("Chart 7").Visible = True

        ws.PivotTables("TCD_return_full").PivotFields( _
        "[FLUVIAL].[Type].[Type]").ClearAllFilters
        ws.PivotTables("TCD_return_full").PivotFields( _
        "[FLUVIAL].[Type].[Type]").CurrentPageName = "[FLUVIAL].[Type].&[2]"
        
    End If
        
     

End Sub

Sub CheckIfReturn()

' Vérifier si la cellule nommée "CheckBox1" est égale à 0
    If Range("CheckBox2").Value = 0 Then
        ' Masquer la colonne F de la feuille "Analyse"
        Sheets("Analyse").Columns("L").Hidden = True
    Else
        ' Afficher la colonne F si la valeur n'est pas 0
        Sheets("Analyse").Columns("L").Hidden = False
    End If
        
End Sub

Sub CheckIfMeasureInValues()

    ' Déclaration des variables
    Dim pt As PivotTable
    Dim cf As CubeField
    Dim ChampExistant As Boolean
    Dim NomChamp As String
    
    ' Définition du nom du champ à vérifier
    NomChamp = "[Measures].[Somme de Cout_tonnes_tout_routier]"
    
    ' Référence au tableau croisé dynamique
    On Error GoTo ErreurPt ' Gestion des erreurs
    Set pt = Worksheets("Analyse").PivotTables("Tableau croisé dynamique1")
    
    ' Confirmation que le TCD a bien été trouvé
    ' Debug.Print "Tableau croisé dynamique trouvé : " & pt.Name
    
    ' Initialisation de la variable
    ChampExistant = False
    
    ' Boucle pour vérifier si la mesure est déjà présente dans les valeurs via CubeFields
    Debug.Print "Nombre de CubeFields dans le TCD : " & pt.CubeFields.Count
    For Each cf In pt.CubeFields
        ' Debug.Print "CubeField trouvé : " & cf.Name
        If cf.Name = NomChamp Then
            If cf.Orientation = xlDataField Then
                ChampExistant = True
                Debug.Print "Le champ est déjà dans les valeurs : " & cf.Name
                Exit For
            End If
        End If
    Next cf
    
    ' Si la mesure n'est pas présente, l'ajouter aux valeurs
    If Not ChampExistant Then
        ' Modification de l'orientation du CubeField pour le placer en tant que champ de valeur
        With pt.CubeFields(NomChamp)
            .Orientation = xlDataField
        End With
        Debug.Print "Le champ a été ajouté aux valeurs : " & NomChamp
    End If

    Exit Sub ' Fin de la procédure si tout est OK
    
ErreurPt:
    Debug.Print "Erreur : Impossible de trouver le TCD ou les champs."
    Debug.Print "Description de l'erreur : " & Err.Description

End Sub
