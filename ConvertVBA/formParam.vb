Public Class formParam
    Private Sub formParam_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtCheminRapportAnadefi.Text = Form1.defCheminRapport
    End Sub

    Private Sub btnAnnuler_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAnnuler.Click
        Me.Dispose()
    End Sub

    Private Sub btnParcourir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParcourir.Click
        Dim ofd As New OpenFileDialog()
        With ofd
            .Filter = "Fichiers Word (*.doc)|*.DOC"
            .InitialDirectory = Environment.SpecialFolder.Desktop
            .Multiselect = False
            .CheckFileExists = True
        End With
        If (ofd.ShowDialog = System.Windows.Forms.DialogResult.OK) Then
            txtCheminRapportAnadefi.Text = ofd.FileName
        End If
        ofd.Dispose()
    End Sub

    Private Sub btnEnregistrerParametres_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnregistrerParametres.Click
        Form1.defCheminRapport = txtCheminRapportAnadefi.Text
        Form1.txtChemin.Text = Form1.defCheminRapport
        Form1.Refresh()
        Me.Dispose()
    End Sub
End Class