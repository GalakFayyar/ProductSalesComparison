<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class VerifMensuelle
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.AideToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AideToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnParcourir = New System.Windows.Forms.Button()
        Me.lblRepertoireCharge = New System.Windows.Forms.Label()
        Me.lblNomFich = New System.Windows.Forms.Label()
        Me.txtNomFichierRecap = New System.Windows.Forms.TextBox()
        Me.progressBar = New System.Windows.Forms.ProgressBar()
        Me.btnExec = New System.Windows.Forms.Button()
        Me.btnQuit = New System.Windows.Forms.Button()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.progressBar2 = New System.Windows.Forms.ProgressBar()
        Me.lblProgress2 = New System.Windows.Forms.Label()
        Me.lblNbFichiers = New System.Windows.Forms.Label()
        Me.checkedListBoxMois = New System.Windows.Forms.CheckedListBox()
        Me.lblEcritureEnCours = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lblXls = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AideToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(608, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'AideToolStripMenuItem
        '
        Me.AideToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AideToolStripMenuItem1})
        Me.AideToolStripMenuItem.Name = "AideToolStripMenuItem"
        Me.AideToolStripMenuItem.Size = New System.Drawing.Size(62, 20)
        Me.AideToolStripMenuItem.Text = "A propos"
        '
        'AideToolStripMenuItem1
        '
        Me.AideToolStripMenuItem1.Name = "AideToolStripMenuItem1"
        Me.AideToolStripMenuItem1.Size = New System.Drawing.Size(106, 22)
        Me.AideToolStripMenuItem1.Text = "Aide"
        '
        'btnParcourir
        '
        Me.btnParcourir.Location = New System.Drawing.Point(9, 22)
        Me.btnParcourir.Name = "btnParcourir"
        Me.btnParcourir.Size = New System.Drawing.Size(75, 23)
        Me.btnParcourir.TabIndex = 1
        Me.btnParcourir.Text = "Parcourir ..."
        Me.btnParcourir.UseVisualStyleBackColor = True
        '
        'lblRepertoireCharge
        '
        Me.lblRepertoireCharge.AutoSize = True
        Me.lblRepertoireCharge.Location = New System.Drawing.Point(90, 27)
        Me.lblRepertoireCharge.Name = "lblRepertoireCharge"
        Me.lblRepertoireCharge.Size = New System.Drawing.Size(138, 13)
        Me.lblRepertoireCharge.TabIndex = 2
        Me.lblRepertoireCharge.Text = "Aucun Repertoire chargé ..."
        '
        'lblNomFich
        '
        Me.lblNomFich.AutoSize = True
        Me.lblNomFich.Location = New System.Drawing.Point(9, 62)
        Me.lblNomFich.Name = "lblNomFich"
        Me.lblNomFich.Size = New System.Drawing.Size(146, 13)
        Me.lblNomFich.TabIndex = 3
        Me.lblNomFich.Text = "Nom du Fichier Récapitulatif :"
        '
        'txtNomFichierRecap
        '
        Me.txtNomFichierRecap.Location = New System.Drawing.Point(161, 59)
        Me.txtNomFichierRecap.Name = "txtNomFichierRecap"
        Me.txtNomFichierRecap.Size = New System.Drawing.Size(175, 20)
        Me.txtNomFichierRecap.TabIndex = 4
        '
        'progressBar
        '
        Me.progressBar.Location = New System.Drawing.Point(9, 89)
        Me.progressBar.Name = "progressBar"
        Me.progressBar.Size = New System.Drawing.Size(347, 10)
        Me.progressBar.TabIndex = 5
        '
        'btnExec
        '
        Me.btnExec.Enabled = False
        Me.btnExec.Location = New System.Drawing.Point(204, 171)
        Me.btnExec.Name = "btnExec"
        Me.btnExec.Size = New System.Drawing.Size(75, 23)
        Me.btnExec.TabIndex = 6
        Me.btnExec.Text = "Comparer"
        Me.btnExec.UseVisualStyleBackColor = True
        '
        'btnQuit
        '
        Me.btnQuit.Location = New System.Drawing.Point(283, 171)
        Me.btnQuit.Name = "btnQuit"
        Me.btnQuit.Size = New System.Drawing.Size(75, 23)
        Me.btnQuit.TabIndex = 7
        Me.btnQuit.Text = "Quitter"
        Me.btnQuit.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.AutoSize = True
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress.Location = New System.Drawing.Point(6, 102)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(22, 13)
        Me.lblProgress.TabIndex = 8
        Me.lblProgress.Text = "0%"
        '
        'progressBar2
        '
        Me.progressBar2.Location = New System.Drawing.Point(8, 136)
        Me.progressBar2.Name = "progressBar2"
        Me.progressBar2.Size = New System.Drawing.Size(348, 10)
        Me.progressBar2.TabIndex = 9
        '
        'lblProgress2
        '
        Me.lblProgress2.AutoSize = True
        Me.lblProgress2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgress2.Location = New System.Drawing.Point(6, 149)
        Me.lblProgress2.Name = "lblProgress2"
        Me.lblProgress2.Size = New System.Drawing.Size(22, 13)
        Me.lblProgress2.TabIndex = 11
        Me.lblProgress2.Text = "0%"
        '
        'lblNbFichiers
        '
        Me.lblNbFichiers.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.lblNbFichiers.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblNbFichiers.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNbFichiers.Location = New System.Drawing.Point(200, 101)
        Me.lblNbFichiers.Name = "lblNbFichiers"
        Me.lblNbFichiers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblNbFichiers.Size = New System.Drawing.Size(156, 14)
        Me.lblNbFichiers.TabIndex = 13
        Me.lblNbFichiers.Text = "aucun"
        Me.lblNbFichiers.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'checkedListBoxMois
        '
        Me.checkedListBoxMois.Enabled = False
        Me.checkedListBoxMois.FormattingEnabled = True
        Me.checkedListBoxMois.Location = New System.Drawing.Point(389, 32)
        Me.checkedListBoxMois.Name = "checkedListBoxMois"
        Me.checkedListBoxMois.Size = New System.Drawing.Size(210, 199)
        Me.checkedListBoxMois.TabIndex = 14
        '
        'lblEcritureEnCours
        '
        Me.lblEcritureEnCours.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.lblEcritureEnCours.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblEcritureEnCours.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEcritureEnCours.Location = New System.Drawing.Point(50, 148)
        Me.lblEcritureEnCours.Name = "lblEcritureEnCours"
        Me.lblEcritureEnCours.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblEcritureEnCours.Size = New System.Drawing.Size(306, 14)
        Me.lblEcritureEnCours.TabIndex = 15
        Me.lblEcritureEnCours.Text = "attente traitement"
        Me.lblEcritureEnCours.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblRepertoireCharge)
        Me.GroupBox1.Controls.Add(Me.lblEcritureEnCours)
        Me.GroupBox1.Controls.Add(Me.btnParcourir)
        Me.GroupBox1.Controls.Add(Me.lblNomFich)
        Me.GroupBox1.Controls.Add(Me.lblNbFichiers)
        Me.GroupBox1.Controls.Add(Me.txtNomFichierRecap)
        Me.GroupBox1.Controls.Add(Me.lblProgress2)
        Me.GroupBox1.Controls.Add(Me.progressBar)
        Me.GroupBox1.Controls.Add(Me.progressBar2)
        Me.GroupBox1.Controls.Add(Me.btnExec)
        Me.GroupBox1.Controls.Add(Me.lblProgress)
        Me.GroupBox1.Controls.Add(Me.btnQuit)
        Me.GroupBox1.Controls.Add(Me.lblXls)
        Me.GroupBox1.Location = New System.Drawing.Point(10, 30)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(373, 201)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Opérations"
        '
        'lblXls
        '
        Me.lblXls.AutoSize = True
        Me.lblXls.Location = New System.Drawing.Point(334, 62)
        Me.lblXls.Margin = New System.Windows.Forms.Padding(0)
        Me.lblXls.Name = "lblXls"
        Me.lblXls.Size = New System.Drawing.Size(22, 13)
        Me.lblXls.TabIndex = 16
        Me.lblXls.Text = ".xls"
        '
        'VerifMensuelle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(608, 242)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.checkedListBoxMois)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "VerifMensuelle"
        Me.Text = "Système de Vérification Mensuelle v.0.1"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents btnParcourir As System.Windows.Forms.Button
    Friend WithEvents lblRepertoireCharge As System.Windows.Forms.Label
    Friend WithEvents lblNomFich As System.Windows.Forms.Label
    Friend WithEvents txtNomFichierRecap As System.Windows.Forms.TextBox
    Friend WithEvents progressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents btnExec As System.Windows.Forms.Button
    Friend WithEvents btnQuit As System.Windows.Forms.Button
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents progressBar2 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblProgress2 As System.Windows.Forms.Label
    Friend WithEvents lblNbFichiers As System.Windows.Forms.Label
    Friend WithEvents AideToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AideToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents checkedListBoxMois As System.Windows.Forms.CheckedListBox
    Friend WithEvents lblEcritureEnCours As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblXls As System.Windows.Forms.Label

End Class
