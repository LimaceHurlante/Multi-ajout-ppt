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
        Me.ButtonPickFile = New System.Windows.Forms.Button()
        Me.ButtonPickFolder = New System.Windows.Forms.Button()
        Me.IncludeSubFolder = New System.Windows.Forms.CheckBox()
        Me.EndSlide = New System.Windows.Forms.RadioButton()
        Me.StartSlide = New System.Windows.Forms.RadioButton()
        Me.ButtonRun = New System.Windows.Forms.Button()
        Me.ButtonExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ButtonPickFile
        '
        Me.ButtonPickFile.Location = New System.Drawing.Point(12, 12)
        Me.ButtonPickFile.Name = "ButtonPickFile"
        Me.ButtonPickFile.Size = New System.Drawing.Size(292, 31)
        Me.ButtonPickFile.TabIndex = 0
        Me.ButtonPickFile.Text = "Choisir le fichier contenant la slide à rajouter"
        Me.ButtonPickFile.UseVisualStyleBackColor = True
        '
        'ButtonPickFolder
        '
        Me.ButtonPickFolder.Location = New System.Drawing.Point(12, 49)
        Me.ButtonPickFolder.Name = "ButtonPickFolder"
        Me.ButtonPickFolder.Size = New System.Drawing.Size(292, 31)
        Me.ButtonPickFolder.TabIndex = 1
        Me.ButtonPickFolder.Text = "Choisir le dossier contenant les ppt à modifier"
        Me.ButtonPickFolder.UseVisualStyleBackColor = True
        '
        'IncludeSubFolder
        '
        Me.IncludeSubFolder.AutoSize = True
        Me.IncludeSubFolder.Checked = True
        Me.IncludeSubFolder.CheckState = System.Windows.Forms.CheckState.Checked
        Me.IncludeSubFolder.Location = New System.Drawing.Point(337, 63)
        Me.IncludeSubFolder.Name = "IncludeSubFolder"
        Me.IncludeSubFolder.Size = New System.Drawing.Size(140, 17)
        Me.IncludeSubFolder.TabIndex = 4
        Me.IncludeSubFolder.Text = "Inclure les sous-dossiers"
        Me.IncludeSubFolder.UseVisualStyleBackColor = True
        '
        'EndSlide
        '
        Me.EndSlide.AutoSize = True
        Me.EndSlide.Checked = True
        Me.EndSlide.Location = New System.Drawing.Point(327, 40)
        Me.EndSlide.Name = "EndSlide"
        Me.EndSlide.Size = New System.Drawing.Size(165, 17)
        Me.EndSlide.TabIndex = 3
        Me.EndSlide.TabStop = True
        Me.EndSlide.Text = "Ajouter la slide à la fin des ppt"
        Me.EndSlide.UseVisualStyleBackColor = True
        '
        'StartSlide
        '
        Me.StartSlide.AutoSize = True
        Me.StartSlide.Location = New System.Drawing.Point(322, 19)
        Me.StartSlide.Name = "StartSlide"
        Me.StartSlide.Size = New System.Drawing.Size(176, 17)
        Me.StartSlide.TabIndex = 2
        Me.StartSlide.Text = "Ajouter la slide au début des ppt"
        Me.StartSlide.UseVisualStyleBackColor = True
        '
        'ButtonRun
        '
        Me.ButtonRun.Location = New System.Drawing.Point(137, 102)
        Me.ButtonRun.Name = "ButtonRun"
        Me.ButtonRun.Size = New System.Drawing.Size(355, 31)
        Me.ButtonRun.TabIndex = 5
        Me.ButtonRun.Text = "Lancer le rajout multiple"
        Me.ButtonRun.UseVisualStyleBackColor = True
        '
        'ButtonExit
        '
        Me.ButtonExit.Location = New System.Drawing.Point(12, 102)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(107, 31)
        Me.ButtonExit.TabIndex = 6
        Me.ButtonExit.Text = "Quitter"
        Me.ButtonExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(494, 139)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(13, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "?"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(513, 155)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ButtonExit)
        Me.Controls.Add(Me.ButtonRun)
        Me.Controls.Add(Me.StartSlide)
        Me.Controls.Add(Me.EndSlide)
        Me.Controls.Add(Me.IncludeSubFolder)
        Me.Controls.Add(Me.ButtonPickFolder)
        Me.Controls.Add(Me.ButtonPickFile)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ButtonPickFile As Button
    Friend WithEvents ButtonPickFolder As Button
    Friend WithEvents IncludeSubFolder As CheckBox
    Friend WithEvents EndSlide As RadioButton
    Friend WithEvents StartSlide As RadioButton
    Friend WithEvents ButtonRun As Button
    Friend WithEvents ButtonExit As Button
    Friend WithEvents Label1 As Label
End Class
