<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class formAide
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
        Me.rchTxtAide = New System.Windows.Forms.RichTextBox
        Me.btnFermer = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'rchTxtAide
        '
        Me.rchTxtAide.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rchTxtAide.Location = New System.Drawing.Point(12, 12)
        Me.rchTxtAide.Name = "rchTxtAide"
        Me.rchTxtAide.ReadOnly = True
        Me.rchTxtAide.Size = New System.Drawing.Size(488, 498)
        Me.rchTxtAide.TabIndex = 0
        Me.rchTxtAide.Text = ""
        '
        'btnFermer
        '
        Me.btnFermer.Location = New System.Drawing.Point(425, 516)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(75, 23)
        Me.btnFermer.TabIndex = 1
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'formAide
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(512, 543)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.rchTxtAide)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "formAide"
        Me.Text = "Aide à l'utilisation"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents rchTxtAide As System.Windows.Forms.RichTextBox
    Friend WithEvents btnFermer As System.Windows.Forms.Button
End Class
