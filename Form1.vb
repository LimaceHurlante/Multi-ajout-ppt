
Imports Graph = Microsoft.Office.Interop.Graph
Imports Office = Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class Form1
    Public FicherPPTdeBASE As String = "", DossierExport As String = "", OnYVa As Boolean, dlgOpen As FileDialog, dlgSave As FileDialog
    Public Shared PwP As PowerPoint.Application
    Public PresDeBase As PowerPoint.Presentation, SlideDeBase As PowerPoint.Slide
    Public PresAModifer As PowerPoint.Presentation, SlideAModifer As PowerPoint.Slide
    'Pour le texte en anglais
    Public SelectedFile As String, NoSelectedFile As String, ChoosePPTFile As String, FilterPPTFile As String, SelectedFolder As String, NoSelectedFolder As String, ChooseFolder As String
    ' TEST LISTE DE FICHIER
    Public Fichiers As New ArrayList()
    Public FromSlideNo As Integer, ToSlideNo As Integer


    'CHARGEMENT DE L'AFFICHAGE
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LoadLanguage()
        Call CheckIfReadyToGo()
        GroupBoxAjoutDeTransition.Left = GroupBoxAjoutDeSlide.Left
        GroupBoxAjoutDeTransition.Top = GroupBoxAjoutDeSlide.Top
        GroupBoxAjoutDeTransition.Visible = False
        LabelAttente.Visible = False
        Me.Width = 325

    End Sub
    Private Sub LoadLanguage()
        If Strings.LCase(Strings.Left(System.Globalization.CultureInfo.InstalledUICulture.Name, 2)) = "fr" Then
            SelectedFile = "Vous avez selectionné ce fichier :"
            NoSelectedFile = "Pas de fichier séléctioné ..."
            ChoosePPTFile = "Choix du fichier PowerPoint à ajouter"
            FilterPPTFile = "Fichiers PowerPoint|*.ppt;*.pptx;*.pptm"
            SelectedFolder = "Vous avez selectionné ce dossier :"
            NoSelectedFolder = "Pas de dossier séléctioné ..."
            ChooseFolder = "Selectionnez le dossier dans lequel se trouvent les ppt à modifer"

        Else
            SelectedFile = "You selected this file:"
            NoSelectedFile = "No selected file ..."
            ChoosePPTFile = "Choose a PowerPoint file to add"
            FilterPPTFile = "PowerPoint files|*.ppt;*.pptx;*.pptm"
            SelectedFolder = "You selected this folder:"
            NoSelectedFolder = "No selected folder ..."
            ChooseFolder = "Select the folder where the ppt to modify are located"
            ButtonPickFile.Text = "Pick the file containing the slide to add"
            ButtonPickFolder.Text = "Pick the ppt to modify"
            ButtonExit.Text = "Quit"
            ButtonRun.Text = "Launch multiple addition"
            StartSlide.Text = "Add the slide at the start of the slide show"
            EndSlide.Text = "Add the slide at the end of the slide show"
            Form2.IncludeSubFolder.Text = "Include subfolders"
        End If
    End Sub

    Private Sub CBBoxDebutACopier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBoxDebutACopier.SelectedIndexChanged
        VerifEtMiseAJourDesVariableDebutEtFinDeSlideACopier()
        MiseAJourLabelNbDeSlidesSingulierPluriel()
    End Sub
    Private Sub CBBoxFinACopier_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBBoxFinACopier.SelectedIndexChanged
        VerifEtMiseAJourDesVariableDebutEtFinDeSlideACopier()
        MiseAJourLabelNbDeSlidesSingulierPluriel()
    End Sub

    Private Sub VerifEtMiseAJourDesVariableDebutEtFinDeSlideACopier()
        If CBBoxFinACopier.SelectedItem < CBBoxDebutACopier.SelectedItem Then
            CBBoxFinACopier.SelectedItem = ToSlideNo
            CBBoxDebutACopier.SelectedItem = FromSlideNo
        Else
            ToSlideNo = CBBoxFinACopier.SelectedItem
            FromSlideNo = CBBoxDebutACopier.SelectedItem
        End If
    End Sub

    Sub MiseAJourLabelNbDeSlidesSingulierPluriel()
        If IsNumeric(CBBoxDebutACopier.SelectedItem) And IsNumeric(CBBoxFinACopier.SelectedItem) Then
            StartSlide.Text = "Ajouter la slide au début des ppt"
            If (CBBoxFinACopier.SelectedItem - CBBoxDebutACopier.SelectedItem + 1) > 1 Then
                StartSlide.Text = "Ajouter les slides au début des ppt"
            Else
                StartSlide.Text = "Ajouter la slide au début des ppt"
            End If
        End If
    End Sub


    Private Sub RadioButtonAjoutTransitions_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonAjoutTransitions.CheckedChanged
        If RadioButtonAjoutTransitions.Checked Then
            GroupBoxAjoutDeSlide.Visible = False
            GroupBoxAjoutDeTransition.Visible = True
        Else
            GroupBoxAjoutDeSlide.Visible = True
            GroupBoxAjoutDeTransition.Visible = False
        End If
    End Sub

    Private Sub ButtonPickFile_Click(sender As Object, e As EventArgs) Handles ButtonPickFile.Click
        FicherPPTdeBASE = FileAddress()
        Dim NbDeSlide As Integer

        Dim NomDuFichierAAfficher As String
        If FicherPPTdeBASE IsNot "" Then
            If Len(FicherPPTdeBASE) < 48 Then
                NomDuFichierAAfficher = FicherPPTdeBASE
            Else
                NomDuFichierAAfficher = "... " & Strings.LTrim(Strings.Right(FicherPPTdeBASE, 45))
            End If
            ButtonPickFile.Text = NomDuFichierAAfficher & vbCrLf & "Pour changer de fichier cliquer ici"
            LabelAttente.Visible = True
            NbDeSlide = NbSlide(FicherPPTdeBASE)
            LabelAttente.Visible = False
            LabelNbSlides.Text = "Ce fichier comprend " & NbDeSlide & " slides."
            FromSlideNo = 1
            ToSlideNo = NbDeSlide
            For i = 1 To NbDeSlide
                CBBoxDebutACopier.Items.Add(i)
                CBBoxFinACopier.Items.Add(i)
            Next
            LabelNbSlides.Enabled = True
            LabelQuellesSlidesCopier.Enabled = True
            CBBoxDebutACopier.Enabled = True
            CBBoxDebutACopier.SelectedIndex = 0
            CBBoxFinACopier.Enabled = True
            CBBoxFinACopier.SelectedIndex = CBBoxFinACopier.Items.Count - 1
            MiseAJourLabelNbDeSlidesSingulierPluriel()

        Else
            MsgBox(NoSelectedFile)
            ButtonPickFile.Text = "Pour choisir de fichier source cliquer ici"
            LabelNbSlides.Text = "Ce fichier comprend  X slides."
            LabelNbSlides.Enabled = False
            LabelQuellesSlidesCopier.Enabled = False
            CBBoxDebutACopier.Enabled = False
            CBBoxFinACopier.Enabled = False
        End If
        Call CheckIfReadyToGo()
    End Sub
    Private Function FileAddress() As String
        Dim ofd As New OpenFileDialog With {
            .Multiselect = False,
            .Title = ChoosePPTFile,
            .Filter = FilterPPTFile
        }

        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return ofd.FileName
        Else
            Return String.Empty
        End If
    End Function
    Private Sub ButtonPickFolder_Click(sender As Object, e As EventArgs) Handles ButtonPickFolder.Click
        Form2.ShowDialog()
        LabelDossierDeTravail.Text = Form2.DossierDeTravail
        LabelNombreDeFichier.Text = Fichiers.Count & " fichiers à modifier"
    End Sub
    Private Function FolderAddress() As String
        Dim dlg As New FolderBrowserDialog With {
            .Description = ChooseFolder
        }
        If dlg.ShowDialog() = DialogResult.OK Then
            Dim path As String = dlg.SelectedPath
            Return path
        Else
            Return String.Empty
        End If
    End Function
    Private Sub ButtonExit_Click(sender As Object, e As EventArgs) Handles ButtonExit.Click
        Application.Exit()
    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        MsgBox("Made by @ZacharieCortes", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Crédits") ', MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
    End Sub
    Function NbSlide(fichierACompter As String) As Integer
        'Start Powerpoint
        PwP = New PowerPoint.Application()
        PresDeBase = PwP.Presentations.Open(fichierACompter, , , False)
        NbSlide = PresDeBase.Slides.Count
        PresDeBase.Saved = True
        PresDeBase.Close()
        PresDeBase = Nothing
        'on quitte pwp
        PwP.Quit()
        PwP = Nothing
        GC.Collect()
    End Function

    Private Sub CheckIfReadyToGo()
        ButtonRun.Enabled = (Not String.IsNullOrEmpty(FicherPPTdeBASE)) And (Not String.IsNullOrEmpty(DossierExport))
        If ButtonRun.Enabled Then ButtonRun.Select()
    End Sub
    Private Sub ButtonRun_Click(sender As Object, e As EventArgs) Handles ButtonRun.Click

        Call Main()

    End Sub
    Sub Main()
        'Start Powerpoint
        PwP = New PowerPoint.Application()

        'copier la slide et refermer le fichier
        PresDeBase = PwP.Presentations.Open(FicherPPTdeBASE, , , False)
        PresDeBase.Slides(1).Copy()
        PresDeBase.Saved = True
        PresDeBase.Close()
        PresDeBase = Nothing

        'boucle sur tout les fichiers
        'Call ListFilesInFolder(DossierExport, Form2.IncludeSubFolder.Checked)

        'on quitte pwp
        PwP.Quit()
        PwP = Nothing
        GC.Collect()
        Me.Close()
    End Sub
    Sub OuvreColleLaSlideSauveEtFerme(nom As String)

        PresAModifer = PwP.Presentations.Open(nom, , , False)
        If StartSlide.Checked Then
            PresAModifer.Slides().Paste(1)
        Else
            PresAModifer.Slides().Paste()
        End If

        PresAModifer.Save()
        PresAModifer.Close()
    End Sub


End Class
