<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtChemin = New System.Windows.Forms.TextBox
        Me.btnTraitementAnadefi = New System.Windows.Forms.Button
        Me.btnParcourir = New System.Windows.Forms.Button
        Me.btnExec = New System.Windows.Forms.Button
        Me.btnFermer = New System.Windows.Forms.Button
        Me.ProgressBar = New System.Windows.Forms.ProgressBar
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
        Me.ParamètresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ParamètresToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.AProposToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.QuitterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.lblProgress = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtChemin)
        Me.GroupBox1.Controls.Add(Me.btnTraitementAnadefi)
        Me.GroupBox1.Controls.Add(Me.btnParcourir)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 27)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(608, 71)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Chemin vers le document"
        '
        'txtChemin
        '
        Me.txtChemin.Location = New System.Drawing.Point(6, 16)
        Me.txtChemin.Name = "txtChemin"
        Me.txtChemin.ReadOnly = True
        Me.txtChemin.Size = New System.Drawing.Size(596, 20)
        Me.txtChemin.TabIndex = 8
        Me.txtChemin.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnTraitementAnadefi
        '
        Me.btnTraitementAnadefi.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTraitementAnadefi.Location = New System.Drawing.Point(296, 42)
        Me.btnTraitementAnadefi.Name = "btnTraitementAnadefi"
        Me.btnTraitementAnadefi.Size = New System.Drawing.Size(225, 23)
        Me.btnTraitementAnadefi.TabIndex = 0
        Me.btnTraitementAnadefi.Text = "Sélectionner Document Anadefi"
        Me.btnTraitementAnadefi.UseVisualStyleBackColor = True
        '
        'btnParcourir
        '
        Me.btnParcourir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnParcourir.Location = New System.Drawing.Point(527, 42)
        Me.btnParcourir.Name = "btnParcourir"
        Me.btnParcourir.Size = New System.Drawing.Size(75, 23)
        Me.btnParcourir.TabIndex = 5
        Me.btnParcourir.Text = "Parcourir ..."
        Me.btnParcourir.UseVisualStyleBackColor = True
        '
        'btnExec
        '
        Me.btnExec.Enabled = False
        Me.btnExec.Location = New System.Drawing.Point(452, 127)
        Me.btnExec.Name = "btnExec"
        Me.btnExec.Size = New System.Drawing.Size(75, 23)
        Me.btnExec.TabIndex = 2
        Me.btnExec.Text = "Exécuter"
        Me.btnExec.UseVisualStyleBackColor = True
        '
        'btnFermer
        '
        Me.btnFermer.Location = New System.Drawing.Point(533, 127)
        Me.btnFermer.Name = "btnFermer"
        Me.btnFermer.Size = New System.Drawing.Size(75, 23)
        Me.btnFermer.TabIndex = 3
        Me.btnFermer.Text = "Fermer"
        Me.btnFermer.UseVisualStyleBackColor = True
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(6, 127)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(440, 22)
        Me.ProgressBar.TabIndex = 4
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ParamètresToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(619, 24)
        Me.MenuStrip1.TabIndex = 7
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ParamètresToolStripMenuItem
        '
        Me.ParamètresToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ParamètresToolStripMenuItem1, Me.AProposToolStripMenuItem, Me.QuitterToolStripMenuItem})
        Me.ParamètresToolStripMenuItem.Name = "ParamètresToolStripMenuItem"
        Me.ParamètresToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.ParamètresToolStripMenuItem.Text = "Outils"
        '
        'ParamètresToolStripMenuItem1
        '
        Me.ParamètresToolStripMenuItem1.Name = "ParamètresToolStripMenuItem1"
        Me.ParamètresToolStripMenuItem1.Size = New System.Drawing.Size(152, 22)
        Me.ParamètresToolStripMenuItem1.Text = "Paramètres"
        '
        'AProposToolStripMenuItem
        '
        Me.AProposToolStripMenuItem.Name = "AProposToolStripMenuItem"
        Me.AProposToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.AProposToolStripMenuItem.Text = "A propos"
        '
        'QuitterToolStripMenuItem
        '
        Me.QuitterToolStripMenuItem.Name = "QuitterToolStripMenuItem"
        Me.QuitterToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.QuitterToolStripMenuItem.Text = "Quitter"
        '
        'lblProgress
        '
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.Location = New System.Drawing.Point(6, 101)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(608, 23)
        Me.lblProgress.TabIndex = 10
        Me.lblProgress.Text = "statutProgressBar"
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblProgress.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(619, 164)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.ProgressBar)
        Me.Controls.Add(Me.btnFermer)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnExec)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "[ANADEFI] Correctif VBA"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnExec As System.Windows.Forms.Button
    Friend WithEvents btnFermer As System.Windows.Forms.Button
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents btnParcourir As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ParamètresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ParamètresToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AProposToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuitterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnTraitementAnadefi As System.Windows.Forms.Button
    Friend WithEvents txtChemin As System.Windows.Forms.TextBox
    Friend WithEvents lblProgress As System.Windows.Forms.Label

End Class
