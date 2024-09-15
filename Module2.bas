Attribute VB_Name = "Module2"
Sub RefreshPowerPivotCharts()

    Dim ws As Worksheet
    Dim chtObj As ChartObject
    Dim pvtTable As PivotTable
    
    ' Rafra�chir le mod�le PowerPivot
    ThisWorkbook.Model.Refresh
    
    ' Boucler sur toutes les feuilles
    For Each ws In ThisWorkbook.Worksheets
        ' Boucler sur tous les TCD de la feuille
        For Each pvtTable In ws.PivotTables
            ' Rafra�chir le TCD
            pvtTable.RefreshTable
        Next pvtTable
        
        ' Boucler sur tous les graphiques de chaque feuille
        For Each chtObj In ws.ChartObjects
            ' Forcer Excel � redessiner le graphique en activant puis d�sactivant la feuille
            ws.Activate
            chtObj.Chart.Refresh
        Next chtObj
    Next ws
    
    CheckIfMeasureInValues
    
    MsgBox "Les graphiques et TCD li�s au mod�le PowerPivot ont �t� mis � jour avec succ�s!", vbInformation
    
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
    
    ' D�sactiver les fonctionnalit�s pour am�liorer les performances
    Application.ScreenUpdating = False ' D�sactiver la mise � jour de l'�cran
    Application.Calculation = xlCalculationManual ' D�sactiver le calcul automatique
    Application.DisplayAlerts = False ' D�sactiver les alertes
    Application.EnableEvents = False ' D�sactiver les �v�nements
    

    ' D�finir la feuille source
    Set wsSource = ThisWorkbook.Sheets("Analyse")
    
    ' Temporiser la feuille Analyse pour l'exportation
    wsSource.Visible = xlSheetVisible ' Rendre l'onglet visible temporairement
    
    ' D�finir la plage nomm�e
    On Error Resume Next
    Set dashboardRange = wsSource.Range("Dashboard")
    On Error GoTo 0
    
    ' V�rifier si la plage nomm�e existe
    If dashboardRange Is Nothing Then
        MsgBox "La plage nomm�e 'Dashboard' n'existe pas sur la feuille 'Analyse'.", vbExclamation
        Exit Sub
    End If
    
    ' Obtenir le chemin du dossier du classeur actif
    currentWorkbookPath = ThisWorkbook.Path
    If currentWorkbookPath = "" Then
        MsgBox "Le classeur doit �tre enregistr� avant de proc�der.", vbExclamation
        Exit Sub
    End If
    
    ' D�finir le chemin du fichier PDF
    pdfFilePath = currentWorkbookPath & "\Dashboard.pdf"
    
    ' D�finir la zone d'impression
    wsSource.PageSetup.PrintArea = dashboardRange.Address
    
    ' Configurer la mise en page en paysage
    With wsSource.PageSetup
        .Orientation = xlLandscape
        .Zoom = False ' Utiliser les dimensions sp�cifi�es pour l'impression
        .FitToPagesWide = 1 ' Ajuster automatiquement la largeur pour ne pas couper en mettant sur False
        .FitToPagesTall = 1
    End With
    
    ' Exporter la plage en PDF
    wsSource.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard
    
    ' Ouvrir le fichier PDF
    Shell "explorer.exe " & pdfFilePath, vbNormalFocus
    
    ' R�activer l'onglet Start
    wsStart.Activate
    
    ' Assurez-vous que l'onglet Analyse reste masqu�
    wsSource.Visible = xlSheetHidden  ' Garder l'onglet Analyse masqu�
    
    ' R�activer les fonctionnalit�s d�sactiv�es
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' Avertir l'utilisateur
    MsgBox "La plage 'Dashboard' a �t� enregistr�e en tant que PDF et le fichier a �t� ouvert.", vbInformation
    
End Sub

Sub FiltrerTCD()

    Dim ws As Worksheet
    
    ' R�f�rence � la feuille contenant le TCD
    Set ws = ThisWorkbook.Sheets("Analyse")
    
    ' Lister tous les graphiques
    ' For Each obj In ws.ChartObjects
        ' Debug.Print "Graphique: " & obj.Name
    ' Next obj
    
    ' V�rifier si la cellule F6 est �gale � 0
    If ws.Range("D7").Value = 0 Then
        ' Masquer le Graphique7
        ws.ChartObjects("Chart 7").Visible = False
    Else
        ' Afficher le Graphique7 et ex�cuter FiltrerTCD
        ws.ChartObjects("Chart 7").Visible = True

        ws.PivotTables("TCD_return_full").PivotFields( _
        "[FLUVIAL].[Type].[Type]").ClearAllFilters
        ws.PivotTables("TCD_return_full").PivotFields( _
        "[FLUVIAL].[Type].[Type]").CurrentPageName = "[FLUVIAL].[Type].&[2]"
        
    End If
        
     

End Sub

Sub CheckIfReturn()

' V�rifier si la cellule nomm�e "CheckBox1" est �gale � 0
    If Range("CheckBox2").Value = 0 Then
        ' Masquer la colonne F de la feuille "Analyse"
        Sheets("Analyse").Columns("L").Hidden = True
    Else
        ' Afficher la colonne F si la valeur n'est pas 0
        Sheets("Analyse").Columns("L").Hidden = False
    End If
        
End Sub

Sub CheckIfMeasureInValues()

    ' D�claration des variables
    Dim pt As PivotTable
    Dim cf As CubeField
    Dim ChampExistant As Boolean
    Dim NomChamp As String
    
    ' D�finition du nom du champ � v�rifier
    NomChamp = "[Measures].[Somme de Cout_tonnes_tout_routier]"
    
    ' R�f�rence au tableau crois� dynamique
    On Error GoTo ErreurPt ' Gestion des erreurs
    Set pt = Worksheets("Analyse").PivotTables("Tableau crois� dynamique1")
    
    ' Confirmation que le TCD a bien �t� trouv�
    ' Debug.Print "Tableau crois� dynamique trouv� : " & pt.Name
    
    ' Initialisation de la variable
    ChampExistant = False
    
    ' Boucle pour v�rifier si la mesure est d�j� pr�sente dans les valeurs via CubeFields
    Debug.Print "Nombre de CubeFields dans le TCD : " & pt.CubeFields.Count
    For Each cf In pt.CubeFields
        ' Debug.Print "CubeField trouv� : " & cf.Name
        If cf.Name = NomChamp Then
            If cf.Orientation = xlDataField Then
                ChampExistant = True
                Debug.Print "Le champ est d�j� dans les valeurs : " & cf.Name
                Exit For
            End If
        End If
    Next cf
    
    ' Si la mesure n'est pas pr�sente, l'ajouter aux valeurs
    If Not ChampExistant Then
        ' Modification de l'orientation du CubeField pour le placer en tant que champ de valeur
        With pt.CubeFields(NomChamp)
            .Orientation = xlDataField
        End With
        Debug.Print "Le champ a �t� ajout� aux valeurs : " & NomChamp
    End If

    Exit Sub ' Fin de la proc�dure si tout est OK
    
ErreurPt:
    Debug.Print "Erreur : Impossible de trouver le TCD ou les champs."
    Debug.Print "Description de l'erreur : " & Err.Description

End Sub
