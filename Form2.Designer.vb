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
        Me.Button1 = New System.Windows.Forms.Button()
        Me.ButtonPickFolder = New System.Windows.Forms.Button()
        Me.CheckedListBox = New System.Windows.Forms.CheckedListBox()
        Me.LabelDirectoryName = New System.Windows.Forms.Label()
        Me.ButtonSelectAll = New System.Windows.Forms.Button()
        Me.ButtonUnselectAll = New System.Windows.Forms.Button()
        Me.LabelNombreDeFichier = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 358)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(118, 29)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Annuler"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ButtonPickFolder
        '
        Me.ButtonPickFolder.Location = New System.Drawing.Point(12, 12)
        Me.ButtonPickFolder.Name = "ButtonPickFolder"
        Me.ButtonPickFolder.Size = New System.Drawing.Size(761, 38)
        Me.ButtonPickFolder.TabIndex = 2
        Me.ButtonPickFolder.Text = "pick Foldeer"
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
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(135, 358)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(638, 29)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "Valider"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(788, 398)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.LabelNombreDeFichier)
        Me.Controls.Add(Me.ButtonUnselectAll)
        Me.Controls.Add(Me.ButtonSelectAll)
        Me.Controls.Add(Me.LabelDirectoryName)
        Me.Controls.Add(Me.CheckedListBox)
        Me.Controls.Add(Me.ButtonPickFolder)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form2"
        Me.Text = "Form2"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents ButtonPickFolder As Button
    Friend WithEvents CheckedListBox As CheckedListBox
    Friend WithEvents LabelDirectoryName As Label
    Friend WithEvents ButtonSelectAll As Button
    Friend WithEvents ButtonUnselectAll As Button
    Friend WithEvents LabelNombreDeFichier As Label
    Friend WithEvents Button2 As Button
End Class
