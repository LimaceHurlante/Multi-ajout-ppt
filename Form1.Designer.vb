<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Me.ButtonPickFile = New System.Windows.Forms.Button()
        Me.ButtonPickFolder = New System.Windows.Forms.Button()
        Me.EndSlide = New System.Windows.Forms.RadioButton()
        Me.StartSlide = New System.Windows.Forms.RadioButton()
        Me.ButtonRun = New System.Windows.Forms.Button()
        Me.ButtonExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LabelDossierDeTravail = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.LabelNombreDeFichier = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBoxAjoutDeSlide = New System.Windows.Forms.GroupBox()
        Me.RadioButtonAjoutSlide = New System.Windows.Forms.RadioButton()
        Me.RadioButtonAjoutTransitions = New System.Windows.Forms.RadioButton()
        Me.GroupBoxAjoutDeTransition = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBoxAjoutDeSlide.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonPickFile
        '
        Me.ButtonPickFile.Location = New System.Drawing.Point(9, 19)
        Me.ButtonPickFile.Name = "ButtonPickFile"
        Me.ButtonPickFile.Size = New System.Drawing.Size(269, 31)
        Me.ButtonPickFile.TabIndex = 0
        Me.ButtonPickFile.Text = "Choisir le fichier contenant la slide à rajouter"
        Me.ButtonPickFile.UseVisualStyleBackColor = True
        '
        'ButtonPickFolder
        '
        Me.ButtonPickFolder.Location = New System.Drawing.Point(5, 21)
        Me.ButtonPickFolder.Name = "ButtonPickFolder"
        Me.ButtonPickFolder.Size = New System.Drawing.Size(275, 31)
        Me.ButtonPickFolder.TabIndex = 1
        Me.ButtonPickFolder.Text = "Choisir les ppt à modifier"
        Me.ButtonPickFolder.UseVisualStyleBackColor = True
        '
        'EndSlide
        '
        Me.EndSlide.AutoSize = True
        Me.EndSlide.Checked = True
        Me.EndSlide.Location = New System.Drawing.Point(14, 116)
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
        Me.StartSlide.Location = New System.Drawing.Point(14, 95)
        Me.StartSlide.Name = "StartSlide"
        Me.StartSlide.Size = New System.Drawing.Size(176, 17)
        Me.StartSlide.TabIndex = 2
        Me.StartSlide.Text = "Ajouter la slide au début des ppt"
        Me.StartSlide.UseVisualStyleBackColor = True
        '
        'ButtonRun
        '
        Me.ButtonRun.Location = New System.Drawing.Point(140, 360)
        Me.ButtonRun.Name = "ButtonRun"
        Me.ButtonRun.Size = New System.Drawing.Size(156, 31)
        Me.ButtonRun.TabIndex = 5
        Me.ButtonRun.Text = "Lancer le rajout multiple"
        Me.ButtonRun.UseVisualStyleBackColor = True
        '
        'ButtonExit
        '
        Me.ButtonExit.Location = New System.Drawing.Point(12, 360)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(121, 31)
        Me.ButtonExit.TabIndex = 6
        Me.ButtonExit.Text = "Quitter"
        Me.ButtonExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(283, 394)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(13, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "?"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 58)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(151, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Dossier contenant les fichiers :"
        '
        'LabelDossierDeTravail
        '
        Me.LabelDossierDeTravail.AutoSize = True
        Me.LabelDossierDeTravail.Location = New System.Drawing.Point(6, 74)
        Me.LabelDossierDeTravail.Name = "LabelDossierDeTravail"
        Me.LabelDossierDeTravail.Size = New System.Drawing.Size(10, 13)
        Me.LabelDossierDeTravail.TabIndex = 9
        Me.LabelDossierDeTravail.Text = "-"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(105, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Fichiers selectionnés"
        '
        'LabelNombreDeFichier
        '
        Me.LabelNombreDeFichier.AutoSize = True
        Me.LabelNombreDeFichier.Location = New System.Drawing.Point(6, 106)
        Me.LabelNombreDeFichier.Name = "LabelNombreDeFichier"
        Me.LabelNombreDeFichier.Size = New System.Drawing.Size(10, 13)
        Me.LabelNombreDeFichier.TabIndex = 11
        Me.LabelNombreDeFichier.Text = "-"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LabelNombreDeFichier)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.LabelDossierDeTravail)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ButtonPickFolder)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(284, 129)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Selection des fichiers"
        '
        'GroupBoxAjoutDeSlide
        '
        Me.GroupBoxAjoutDeSlide.Controls.Add(Me.StartSlide)
        Me.GroupBoxAjoutDeSlide.Controls.Add(Me.EndSlide)
        Me.GroupBoxAjoutDeSlide.Controls.Add(Me.ButtonPickFile)
        Me.GroupBoxAjoutDeSlide.Location = New System.Drawing.Point(12, 174)
        Me.GroupBoxAjoutDeSlide.Name = "GroupBoxAjoutDeSlide"
        Me.GroupBoxAjoutDeSlide.Size = New System.Drawing.Size(284, 172)
        Me.GroupBoxAjoutDeSlide.TabIndex = 13
        Me.GroupBoxAjoutDeSlide.TabStop = False
        Me.GroupBoxAjoutDeSlide.Text = "Ajout de slides"
        '
        'RadioButtonAjoutSlide
        '
        Me.RadioButtonAjoutSlide.AutoSize = True
        Me.RadioButtonAjoutSlide.Checked = True
        Me.RadioButtonAjoutSlide.Location = New System.Drawing.Point(17, 151)
        Me.RadioButtonAjoutSlide.Name = "RadioButtonAjoutSlide"
        Me.RadioButtonAjoutSlide.Size = New System.Drawing.Size(93, 17)
        Me.RadioButtonAjoutSlide.TabIndex = 14
        Me.RadioButtonAjoutSlide.TabStop = True
        Me.RadioButtonAjoutSlide.Text = "Ajout de slides"
        Me.RadioButtonAjoutSlide.UseVisualStyleBackColor = True
        '
        'RadioButtonAjoutTransitions
        '
        Me.RadioButtonAjoutTransitions.AutoSize = True
        Me.RadioButtonAjoutTransitions.Location = New System.Drawing.Point(155, 151)
        Me.RadioButtonAjoutTransitions.Name = "RadioButtonAjoutTransitions"
        Me.RadioButtonAjoutTransitions.Size = New System.Drawing.Size(114, 17)
        Me.RadioButtonAjoutTransitions.TabIndex = 15
        Me.RadioButtonAjoutTransitions.Text = "Ajout de transitions"
        Me.RadioButtonAjoutTransitions.UseVisualStyleBackColor = True
        '
        'GroupBoxAjoutDeTransition
        '
        Me.GroupBoxAjoutDeTransition.Location = New System.Drawing.Point(302, 174)
        Me.GroupBoxAjoutDeTransition.Name = "GroupBoxAjoutDeTransition"
        Me.GroupBoxAjoutDeTransition.Size = New System.Drawing.Size(294, 172)
        Me.GroupBoxAjoutDeTransition.TabIndex = 14
        Me.GroupBoxAjoutDeTransition.TabStop = False
        Me.GroupBoxAjoutDeTransition.Text = "Ajout de transitions"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(648, 411)
        Me.Controls.Add(Me.GroupBoxAjoutDeTransition)
        Me.Controls.Add(Me.RadioButtonAjoutTransitions)
        Me.Controls.Add(Me.RadioButtonAjoutSlide)
        Me.Controls.Add(Me.GroupBoxAjoutDeSlide)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ButtonExit)
        Me.Controls.Add(Me.ButtonRun)
        Me.Name = "Form1"
        Me.Text = "Multi Ajout PPT"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBoxAjoutDeSlide.ResumeLayout(False)
        Me.GroupBoxAjoutDeSlide.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ButtonPickFile As Button
    Friend WithEvents ButtonPickFolder As Button
    Friend WithEvents EndSlide As RadioButton
    Friend WithEvents StartSlide As RadioButton
    Friend WithEvents ButtonRun As Button
    Friend WithEvents ButtonExit As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents LabelDossierDeTravail As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents LabelNombreDeFichier As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBoxAjoutDeSlide As GroupBox
    Friend WithEvents RadioButtonAjoutSlide As RadioButton
    Friend WithEvents RadioButtonAjoutTransitions As RadioButton
    Friend WithEvents GroupBoxAjoutDeTransition As GroupBox
End Class
