<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class formParam
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnParcourir = New System.Windows.Forms.Button
        Me.txtCheminRapportAnadefi = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnEnregistrerParametres = New System.Windows.Forms.Button
        Me.btnAnnuler = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnParcourir
        '
        Me.btnParcourir.Location = New System.Drawing.Point(384, 51)
        Me.btnParcourir.Name = "btnParcourir"
        Me.btnParcourir.Size = New System.Drawing.Size(75, 23)
        Me.btnParcourir.TabIndex = 0
        Me.btnParcourir.Text = "Parcourir"
        Me.btnParcourir.UseVisualStyleBackColor = True
        '
        'txtCheminRapportAnadefi
        '
        Me.txtCheminRapportAnadefi.Location = New System.Drawing.Point(12, 25)
        Me.txtCheminRapportAnadefi.Name = "txtCheminRapportAnadefi"
        Me.txtCheminRapportAnadefi.Size = New System.Drawing.Size(447, 20)
        Me.txtCheminRapportAnadefi.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(223, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Chemin vers le Rapport ANADEFI par défaut :"
        '
        'btnEnregistrerParametres
        '
        Me.btnEnregistrerParametres.Location = New System.Drawing.Point(303, 80)
        Me.btnEnregistrerParametres.Name = "btnEnregistrerParametres"
        Me.btnEnregistrerParametres.Size = New System.Drawing.Size(75, 23)
        Me.btnEnregistrerParametres.TabIndex = 3
        Me.btnEnregistrerParametres.Text = "Valider"
        Me.btnEnregistrerParametres.UseVisualStyleBackColor = True
        '
        'btnAnnuler
        '
        Me.btnAnnuler.Location = New System.Drawing.Point(384, 80)
        Me.btnAnnuler.Name = "btnAnnuler"
        Me.btnAnnuler.Size = New System.Drawing.Size(75, 23)
        Me.btnAnnuler.TabIndex = 4
        Me.btnAnnuler.Text = "Annuler"
        Me.btnAnnuler.UseVisualStyleBackColor = True
        '
        'formParam
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(471, 116)
        Me.Controls.Add(Me.btnAnnuler)
        Me.Controls.Add(Me.btnEnregistrerParametres)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtCheminRapportAnadefi)
        Me.Controls.Add(Me.btnParcourir)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "formParam"
        Me.Text = "Paramétrage Rapports ANADEFI"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnParcourir As System.Windows.Forms.Button
    Friend WithEvents txtCheminRapportAnadefi As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnEnregistrerParametres As System.Windows.Forms.Button
    Friend WithEvents btnAnnuler As System.Windows.Forms.Button
End Class
