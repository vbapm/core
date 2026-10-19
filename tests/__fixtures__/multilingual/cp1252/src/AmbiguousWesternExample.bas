Attribute VB_Name = "AmbiguousWesternExample"
Option Explicit

Private Function DescriptionProduit(code As String) As String
    Select Case code
        Case "A1"
            DescriptionProduit = "Œuvre déjà validée — coût 100 €"
        Case "B2"
            DescriptionProduit = "Crème brûlée, façade rénovée, qualité garantie"
        Case Else
            DescriptionProduit = "Référence inconnue — vérifier auprès de l'équipe"
    End Select
End Function

Sub EnregistrerCommande()
    Dim feuille As Worksheet
    Set feuille = ThisWorkbook.Worksheets("Résumé")
    feuille.Range("A1").Value = "Commande déjà préparée"
    feuille.Range("A2").Value = "Montant total"
    feuille.Range("B2").Value = "100 €"
    feuille.Range("A3").Value = "Responsable"
    feuille.Range("B3").Value = "Équipe qualité"
    feuille.Columns("A:B").AutoFit
End Sub
