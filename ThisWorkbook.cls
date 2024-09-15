VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
        ' Appeler la fonction DeleteRows
    DeleteRows
    
    ' Enregistrer le fichier
    Me.Save
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    ' Vérifier si la feuille désactivée est "Input"
    If Sh.Name = "Input" Then
        ' Cacher la feuille "Input"
        Sh.Visible = xlSheetHidden
        
        ' Réinitialiser le bouton Toggle sur la feuille "Start"
        Sheets("Start").ToggleButton1.Value = False
    End If
End Sub





