
Imports Graph = Microsoft.Office.Interop.Graph
Imports Office = Microsoft.Office.Core
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint

Public Class Form1
    Public FicherPPTdeBASE As String = "", Fichiers As New ArrayList(), FromSlideNo As Integer, ToSlideNo As Integer
    Public Shared PwP As PowerPoint.Application
    'Pour le texte en anglais
    Public NoSelectedFile As String, ChoosePPTFile As String, FilterPPTFile As String, NoSelectedFolder As String, ChooseFolder As String, HowManySlide As String, FileSelected As String

    'TO DO
    '-Faire un decompte des fichiers fait durant l'execution
    '-Bug si dossier trop grand selectionné !

    'CHARGEMENT DE L'AFFICHAGE
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LoadLanguage()
        Call CheckIfReadyToGo()
        GroupBoxAjoutDeTransition.Left = GroupBoxAjoutDeSlide.Left
        GroupBoxAjoutDeTransition.Top = GroupBoxAjoutDeSlide.Top
        GroupBoxAjoutDeTransition.Visible = False
        LabelAttente.Visible = False
        LabelAttenteMain.Visible = False
        Me.Width = 325

    End Sub
    Private Sub LoadLanguage()
        If Strings.LCase(Strings.Left(System.Globalization.CultureInfo.InstalledUICulture.Name, 2)) = "fr" Then
            'VARIABLES
            NoSelectedFile = "Pas de fichier séléctioné ..."
            ChoosePPTFile = "Choix du fichier PowerPoint à ajouter"
            FilterPPTFile = "Fichiers PowerPoint|*.ppt;*.pptx;*.pptm"
            NoSelectedFolder = "Pas de dossier séléctioné ..."
            ChooseFolder = "Selectionnez le dossier dans lequel se trouvent les ppt à modifer"
            HowManySlide = "Ce fichier contient"
            FileSelected = "Nombre de fichiers selectionné(s) : "

        Else
            'VARIABLES
            NoSelectedFile = "No selected file ..."
            ChoosePPTFile = "Choose a PowerPoint file to add"
            FilterPPTFile = "PowerPoint files|*.ppt;*.pptx;*.pptm"
            NoSelectedFolder = "No selected folder ..."
            ChooseFolder = "Select the folder where the ppt to modify are located"
            HowManySlide = "This file contains"
            FileSelected = "Files selected:  "

            'Text des objects
            GroupBoxChoixDuDossier.Text = "File selection"
            LabelShowFolderName.Text = "Folder containing files:"
            LabelShowSelectedFile.Text = "Selected files:"
            RadioButtonAjoutSlide.Text = "Add slides"
            RadioButtonAjoutTransitions.Text = "add transitions"
            GroupBoxAjoutDeSlide.Text = "Adding slides"
            GroupBoxAjoutDeTransition.Text = "Ading transtions"
            LabelAttenteMain.Text = "Work in progress. Please wait"
            LabelAttente.Text = "Wait while processing the source file"
            ButtonPickFile.Text = "Pick the file containing the slide to add"
            ButtonPickFolder.Text = "Pick the ppt to modify"
            ButtonExit.Text = "Quit"
            ButtonRun.Text = "Launch multiple addition"
            StartSlide.Text = "Add the slide at the start of the slide show"
            EndSlide.Text = "Add the slide at the end of the slide show"
            LabelTBC.Text = "Coming soon ..."
            LabelQuellesSlidesCopier.Text = "Add  all  slides  from                to"

            'FORM2
            With Form2
                .IncludeSubFolder.Text = "Include subfolders"
                .Text = "File selection"
                .ButtonPickFolder.Text = "Pick the folfer containing the files"
                .ButtonSelectAll.Text = "Select all"
                .ButtonUnselectAll.Text = "Unselect all"
                .ButtonCancel.Text = "Cancel"
                .ButtonValider.Text = "Validate"
                .LabelNombreDeFichier.Text = "Files selected: "
            End With


        End If
    End Sub

    'CHOIX DES FICHIERS A MODIFIERS
    Private Sub ButtonPickFolder_Click(sender As Object, e As EventArgs) Handles ButtonPickFolder.Click
        Form2.ShowDialog()
        LabelDossierDeTravail.Text = Form2.DossierDeTravail
        LabelNombreDeFichier.Text = Fichiers.Count & " fichiers à modifier"
    End Sub


    'AFFICHAGE AJOUT DE SLIDE OU DE TRANSITIONS
    Private Sub RadioButtonAjoutTransitions_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButtonAjoutTransitions.CheckedChanged
        If RadioButtonAjoutTransitions.Checked Then
            GroupBoxAjoutDeSlide.Visible = False
            GroupBoxAjoutDeTransition.Visible = True
        Else
            GroupBoxAjoutDeSlide.Visible = True
            GroupBoxAjoutDeTransition.Visible = False
        End If
    End Sub

    'CHOIX DU FICHIER SOURCE A AJOUTER
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
            LabelNbSlides.Text = HowManySlide & " " & NbDeSlide & " slides."
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
            LabelNbSlides.Text = HowManySlide & " X slides."
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
    Sub MiseAJourLabelNbDeSlidesSingulierPluriel()
        If IsNumeric(CBBoxDebutACopier.SelectedItem) And IsNumeric(CBBoxFinACopier.SelectedItem) Then
            StartSlide.Text = "Ajouter la slide au début des ppt"
            If (CBBoxFinACopier.SelectedItem - CBBoxDebutACopier.SelectedItem + 1) > 1 Then
                StartSlide.Text = "Ajouter les slides au début des ppt"
                EndSlide.Text = "Ajouter les slides à la fin des ppt"
            Else
                StartSlide.Text = "Ajouter la slide au début des ppt"
                EndSlide.Text = "Ajouter la slide à la fin des ppt"
            End If
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


    Function NbSlide(fichierACompter As String) As Integer
        'Start Powerpoint
        Dim PresDeBase As PowerPoint.Presentation
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

    'BOUTON QUITTER ET CREDIT
    Private Sub ButtonExit_Click(sender As Object, e As EventArgs) Handles ButtonExit.Click
        Application.Exit()
    End Sub
    Private Sub LabelCredits_Click(sender As Object, e As EventArgs) Handles LabelCredits.Click
        MsgBox("Made by @ZacharieCortes", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Crédits") ', MsgBoxStyle.OkOnly Or MsgBoxStyle.Information)
    End Sub

    'BOUTON RUN
    Private Sub CheckIfReadyToGo()
        ButtonRun.Enabled = (Not String.IsNullOrEmpty(FicherPPTdeBASE)) And (Fichiers.Count > 0)
        If ButtonRun.Enabled Then ButtonRun.Select()
    End Sub
    Private Sub ButtonRun_Click(sender As Object, e As EventArgs) Handles ButtonRun.Click
        Call Main()
    End Sub

    'AJOUT  DES SLIDES
    Sub Main()
        Dim Fichier As String
        GroupBoxAjoutDeSlide.Visible = False
        GroupBoxAjoutDeTransition.Visible = False
        LabelAttenteMain.Visible = True
        PwP = New PowerPoint.Application
        For Each Fichier In Fichiers
            Call OuvreAjouteSauveEtFerme(Fichier)
        Next Fichier
        PwP.Quit()
        PwP = Nothing
        GC.Collect()
        LabelAttenteMain.Visible = False
        GroupBoxAjoutDeSlide.Visible = True
        If MsgBox("Vos fichiers ont bien été modifié, souhaitez vous ouvrir le dossier et quitter ?", MsgBoxStyle.YesNo) = 6 Then
            Process.Start("explorer.exe", Form2.DossierDeTravail)
            Application.Exit()
        End If
    End Sub
    Sub OuvreAjouteSauveEtFerme(Fichier As String)
        Dim PptAModifer As PowerPoint.Presentation
        Dim ApresQuelleSlide As Integer = 0
        PptAModifer = PwP.Presentations.Open(FileName:=Fichier, , , False)
        If EndSlide.Checked Then ApresQuelleSlide = PptAModifer.Slides.Count
        'expression.InsertFromFile(CheminPresentationSource, Index, SlideDebut, SlideFin)
        PptAModifer.Slides.InsertFromFile(FicherPPTdeBASE, ApresQuelleSlide, FromSlideNo, ToSlideNo)
        PptAModifer.Save()
        PptAModifer.Close()
    End Sub
End Class
