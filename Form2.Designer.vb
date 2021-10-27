<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.ButtonCancel = New System.Windows.Forms.Button()
        Me.ButtonPickFolder = New System.Windows.Forms.Button()
        Me.CheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.LabelDirectoryName = New System.Windows.Forms.Label()
        Me.ButtonSelectAll = New System.Windows.Forms.Button()
        Me.ButtonUnselectAll = New System.Windows.Forms.Button()
        Me.LabelNombreDeFichier = New System.Windows.Forms.Label()
        Me.ButtonValider = New System.Windows.Forms.Button()
        Me.IncludeSubFolder = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(12, 358)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(118, 29)
        Me.ButtonCancel.TabIndex = 0
        Me.ButtonCancel.Text = "Annuler"
        Me.ButtonCancel.UseVisualStyleBackColor = True
        '
        'ButtonPickFolder
        '
        Me.ButtonPickFolder.Location = New System.Drawing.Point(12, 12)
        Me.ButtonPickFolder.Name = "ButtonPickFolder"
        Me.ButtonPickFolder.Size = New System.Drawing.Size(615, 38)
        Me.ButtonPickFolder.TabIndex = 2
        Me.ButtonPickFolder.Text = "Choisir le dossier dans lequel se trouvent les ppt"
        Me.ButtonPickFolder.UseVisualStyleBackColor = True
        '
        'CheckedListBox
        '
        Me.CheckedListBox.FormattingEnabled = True
        Me.CheckedListBox.Location = New System.Drawing.Point(13, 108)
        Me.CheckedListBox.Name = "CheckedListBox"
        Me.CheckedListBox.Size = New System.Drawing.Size(760, 244)
        Me.CheckedListBox.TabIndex = 3
        '
        'LabelDirectoryName
        '
        Me.LabelDirectoryName.AutoSize = True
        Me.LabelDirectoryName.Location = New System.Drawing.Point(15, 57)
        Me.LabelDirectoryName.Name = "LabelDirectoryName"
        Me.LabelDirectoryName.Size = New System.Drawing.Size(115, 13)
        Me.LabelDirectoryName.TabIndex = 4
        Me.LabelDirectoryName.Text = "__________________"
        '
        'ButtonSelectAll
        '
        Me.ButtonSelectAll.Enabled = False
        Me.ButtonSelectAll.Location = New System.Drawing.Point(12, 75)
        Me.ButtonSelectAll.Name = "ButtonSelectAll"
        Me.ButtonSelectAll.Size = New System.Drawing.Size(118, 27)
        Me.ButtonSelectAll.TabIndex = 5
        Me.ButtonSelectAll.Text = "Selectionner tout"
        Me.ButtonSelectAll.UseVisualStyleBackColor = True
        '
        'ButtonUnselectAll
        '
        Me.ButtonUnselectAll.Enabled = False
        Me.ButtonUnselectAll.Location = New System.Drawing.Point(135, 75)
        Me.ButtonUnselectAll.Name = "ButtonUnselectAll"
        Me.ButtonUnselectAll.Size = New System.Drawing.Size(120, 27)
        Me.ButtonUnselectAll.TabIndex = 6
        Me.ButtonUnselectAll.Text = "Deselectionner tout"
        Me.ButtonUnselectAll.UseVisualStyleBackColor = True
        '
        'LabelNombreDeFichier
        '
        Me.LabelNombreDeFichier.AutoSize = True
        Me.LabelNombreDeFichier.Location = New System.Drawing.Point(261, 82)
        Me.LabelNombreDeFichier.Name = "LabelNombreDeFichier"
        Me.LabelNombreDeFichier.Size = New System.Drawing.Size(172, 13)
        Me.LabelNombreDeFichier.TabIndex = 7
        Me.LabelNombreDeFichier.Text = "Nombre de fichiers selectionné(s) : "
        '
        'ButtonValider
        '
        Me.ButtonValider.Location = New System.Drawing.Point(135, 358)
        Me.ButtonValider.Name = "ButtonValider"
        Me.ButtonValider.Size = New System.Drawing.Size(638, 29)
        Me.ButtonValider.TabIndex = 8
        Me.ButtonValider.Text = "Valider"
        Me.ButtonValider.UseVisualStyleBackColor = True
        '
        'IncludeSubFolder
        '
        Me.IncludeSubFolder.AutoSize = True
        Me.IncludeSubFolder.Checked = True
        Me.IncludeSubFolder.CheckState = System.Windows.Forms.CheckState.Checked
        Me.IncludeSubFolder.Location = New System.Drawing.Point(633, 24)
        Me.IncludeSubFolder.Name = "IncludeSubFolder"
        Me.IncludeSubFolder.Size = New System.Drawing.Size(140, 17)
        Me.IncludeSubFolder.TabIndex = 9
        Me.IncludeSubFolder.Text = "Inclure les sous-dossiers"
        Me.IncludeSubFolder.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(788, 398)
        Me.Controls.Add(Me.IncludeSubFolder)
        Me.Controls.Add(Me.ButtonValider)
        Me.Controls.Add(Me.LabelNombreDeFichier)
        Me.Controls.Add(Me.ButtonUnselectAll)
        Me.Controls.Add(Me.ButtonSelectAll)
        Me.Controls.Add(Me.LabelDirectoryName)
        Me.Controls.Add(Me.CheckedListBox)
        Me.Controls.Add(Me.ButtonPickFolder)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Name = "Form2"
        Me.Text = "Choix des fichiers"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ButtonCancel As Button
    Friend WithEvents ButtonPickFolder As Button
    Friend WithEvents CheckedListBox As CheckedListBox
    Friend WithEvents LabelDirectoryName As Label
    Friend WithEvents ButtonSelectAll As Button
    Friend WithEvents ButtonUnselectAll As Button
    Friend WithEvents LabelNombreDeFichier As Label
    Friend WithEvents ButtonValider As Button
    Friend WithEvents IncludeSubFolder As CheckBox
End Class
