Public Class formAide

    Private Sub formAide_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        rchTxtAide.Text = "Aide à l'utilisation de l'outil" & vbNewLine & _
                        "Le but de cet outil est de permettre la mise à jour des fonctionnalités applicatives " & _
                        "(c'est à dire les macros) afin de rendre le document utilisable avec Office Word 2003 et " & _
                        "Office Word 2010." & vbNewLine & _
                        "Cette section décrit comment utiliser l'application en présntant les contrôles élémentaires" & _
                        vbNewLine & _
                        vbNewLine & _
                        "La page principale est composée d'un champ texte non saisissable et de 3 boutons aux opérations " & _
                        "distinctes :" & _
                        vbNewLine & vbTab & "- Le champ texte non saissable présente le chemin vers le document Anadefi, sur " & _
                        "le poste local." & _
                        vbNewLine & vbTab & "- Le bouton ""Sélectionner Document Anadefi"" permet de charger dans l'outil, " & _
                        "le chemin par défaut vers le rapport." & _
                        vbNewLine & vbTab & "- Le bouton ""Parcourir"" permet de rechercher un rapport sur le poste de travail " & _
                        "et de mettre à jour ce fichier." & _
                        vbNewLine & vbTab & "- Le bouton ""Exécuter"" permet de lancer le traitement de mise à jour du fichier." & _
                        vbNewLine & vbTab & "- Le bouton ""Fermer"" permet de quitter l'outil." & _
                        vbNewLine & _
                        vbNewLine & _
                        vbNewLine & "Procédure d'utilisation : " & _
                        vbNewLine & "Cet outil nécessite qu'un dossier Anadefi soit ouvert et en cours d'utilisation. De plus, " & _
                        "un rapport doit être en cours de consultation sur le poste (la consultation signifie que le rapport est " & _
                        "récupéré sur le PC). Suivre les étapes ci-dessous pour utiliser convenablement l'outil :" & _
                        vbNewLine & "- Ouvrir un dossier client via Anadefi" & _
                        vbNewLine & "- Ouvrir un rapport" & _
                        vbNewLine & "- Fermer le rapport sans enregistrer les modifications" & _
                        vbNewLine & "- Lancer l'outil" & _
                        vbNewLine & "- Vérifier que le chemin ""Chemin vers le document"" est renseigné. S'il n'est pas renseigné, " & _
                        "cliquer sur ""Sélectionner Document Anadefi"" pur charger le chemin vers le rapport" & _
                        vbNewLine & "- Cliquer sur ""Exécuter"" pour lancer le traitement" & _
                        vbNewLine & _
                        vbNewLine & "Une fois le message ""Traitement terminé"" affiché, le rapport a été modifié. Utiliser Anadefi " & _
                        "pour enregistrer le document modifié sur le serveur."

    End Sub

    Private Sub btnFermer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFermer.Click
        Me.Dispose()
    End Sub
End Class