
Imports Office = Microsoft.Office.Core
Imports Graph = Microsoft.Office.Interop.Graph
Imports PowerPoint = Microsoft.Office.Interop.PowerPoint
Public Class Form1
    Public FicherPPTdeBASE As String = "", DossierExport As String = "", OnYVa As Boolean, dlgOpen As FileDialog, dlgSave As FileDialog
    Public Shared PwP As PowerPoint.Application
    Public PresDeBase As PowerPoint.Presentation, SlideDeBase As PowerPoint.Slide
    Public PresAModifer As PowerPoint.Presentation, SlideAModifer As PowerPoint.Slide
    'Pour le texte en anglais
    Public SelectedFile As String, NoSelectedFile As String, ChoosePPTFile As String, FilterPPTFile As String, SelectedFolder As String, NoSelectedFolder As String, ChooseFolder As String

    Private Sub CheckIfReadyToGo()
        ButtonRun.Enabled = (Not String.IsNullOrEmpty(FicherPPTdeBASE)) And (Not String.IsNullOrEmpty(DossierExport))
        If ButtonRun.Enabled Then ButtonRun.Select()
    End Sub
    Private Sub ButtonRun_Click(sender As Object, e As EventArgs) Handles ButtonRun.Click

        Call main()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call LoadLanguage()
        Call CheckIfReadyToGo()
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
            ButtonPickFolder.Text = "Pick the folder containing the ppt to modify"
            ButtonExit.Text = "Quit"
            ButtonRun.Text = "Launch multiple addition"
            StartSlide.Text = "Add the slide at the start of the slide show"
            EndSlide.Text = "Add the slide at the end of the slide show"
            IncludeSubFolder.Text = "Include subfolders"
        End If
    End Sub
    Private Sub ButtonPickFile_Click(sender As Object, e As EventArgs) Handles ButtonPickFile.Click
        FicherPPTdeBASE = FileAddress()
        If FicherPPTdeBASE IsNot "" Then
            MsgBox(SelectedFile & vbCrLf & FicherPPTdeBASE)
        Else
            MsgBox(NoSelectedFile)
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
        DossierExport = FolderAddress()
        If DossierExport IsNot "" Then
            MsgBox(SelectedFolder & vbCrLf & DossierExport)
        Else
            MsgBox(NoSelectedFolder)
        End If
        Call CheckIfReadyToGo()
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
        Call ListFilesInFolder(DossierExport, IncludeSubFolder.Checked)

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

    Sub ListFilesInFolder(ByVal xFolderName As String, ByVal xIsSubfolders As Boolean)
        Dim xFileSystemObject As Object
        Dim xFolder As Object
        Dim xSubFolder As Object
        Dim xFile As Object
        xFileSystemObject = CreateObject("Scripting.FileSystemObject")
        xFolder = xFileSystemObject.GetFolder(xFolderName)

        For Each xFile In xFolder.Files
            Dim ext = Microsoft.VisualBasic.Right(xFile.name, 4)
            If ext = ".ppt" Or ext = "pptx" Or ext = "pptm" Or ext = "ppsm" Then
                Debug.Print(xFile.name)
                Call OuvreColleLaSlideSauveEtFerme(xFile.path)
            End If
            'Call ajoutDeLaSlide(xFile)
        Next xFile
        If xIsSubfolders Then
            For Each xSubFolder In xFolder.SubFolders
                Call ListFilesInFolder(xSubFolder.Path, True)
            Next xSubFolder
        End If
        xFile = Nothing
        xFolder = Nothing
        xFileSystemObject = Nothing
    End Sub



    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)
        MsgBox(GetOSLanguage())

    End Sub
    Public Function GetOSLanguage() As String
        Return System.Globalization.CultureInfo.InstalledUICulture.Name
    End Function

End Class



'Sub ARCHIVE_OuvrirLeFichier(leFichier As String)
'    Const sPic = "c:\Program Files (x86)\Microsoft Office\root\CLIPART\PUB60COR\PH02746U.BMP"

'    Dim PwP As PowerPoint.Application
'    Dim oPres As PowerPoint.Presentation
'    Dim oSlide As PowerPoint.Slide

'    'Start Powerpoint and make its window visible but minimized.
'    PwP = New PowerPoint.Application()
'    PwP.Visible = True
'    PwP.WindowState = PowerPoint.PpWindowState.ppWindowMinimized

'    'oPres.Open(FileName:="D:\Dropbox\Boulot\MaMa Xls to PPT\2021\CREDIT TEST.pptx")
'    'PwP.Presentations.Open(FicherPPTdeBASE)
'    Dim sTemplate As String
'    sTemplate = leFichier
'    'Const sTemplate = leFichier
'    'Create a new presentation based on the specified template.
'    oPres = PwP.Presentations.Open(sTemplate, , , True)

'    'Build Slide #1:
'    'Add text to the slide, change the font and insert/position a 
'    'picture on the first slide.
'    oSlide = oPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly)
'    With oSlide.Shapes.Item(1).TextFrame.TextRange
'        .Text = "My Sample Presentation"
'        .Font.Name = "Comic Sans MS"
'        .Font.Size = 48
'    End With
'    oSlide.Shapes.AddPicture(sPic, False, True, 150, 150, 500, 350)
'    oSlide = Nothing

'    'Build Slide #2:
'    'Add text to the slide title, format the text. Also add a chart to the
'    'slide and change the chart type to a 3D pie chart.
'    oSlide = oPres.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutTitleOnly)
'    With oSlide.Shapes.Item(1).TextFrame.TextRange
'        .Text = "My Chart"
'        .Font.Name = "Comic Sans MS"
'        .Font.Size = 48
'    End With
'    Dim oChart As Graph.Chart
'    oChart = oSlide.Shapes.AddOLEObject(150, 150, 480, 320,
'                "MSGraph.Chart.8").OLEFormat.Object
'    oChart.ChartType = Graph.XlChartType.xl3DPie
'    oChart = Nothing
'    oSlide = Nothing

'    'Build Slide #3:
'    'Add a text effect to the slide and apply shadows to the text effect.
'    oSlide = oPres.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutBlank)
'    oSlide.FollowMasterBackground = False
'    Dim oShape As PowerPoint.Shape
'    oShape = oSlide.Shapes.AddTextEffect(Office.MsoPresetTextEffect.msoTextEffect27,
'        "The End", "Impact", 96, False, False, 230, 200)
'    oShape.Shadow.ForeColor.SchemeColor = PowerPoint.PpColorSchemeIndex.ppForeground
'    oShape.Shadow.Visible = True
'    oShape.Shadow.OffsetX = 3
'    oShape.Shadow.OffsetY = 3
'    oShape = Nothing
'    oSlide = Nothing

'    'Modify the slide show transition settings for all 3 slides in
'    'the presentation.
'    Dim SlideIdx(3) As Integer
'    SlideIdx(0) = 1
'    SlideIdx(1) = 2
'    SlideIdx(2) = 3
'    With oPres.Slides.Range(SlideIdx).SlideShowTransition
'        .AdvanceOnTime = True
'        .AdvanceTime = 3
'        .EntryEffect = PowerPoint.PpEntryEffect.ppEffectBoxOut
'    End With
'    Dim oSettings As PowerPoint.SlideShowSettings
'    oSettings = oPres.SlideShowSettings
'    oSettings.StartingSlide = 1
'    oSettings.EndingSlide = 3



'    'Close the presentation without saving changes and quit PowerPoint.
'    oPres.Saved = False
'    oPres.Save()
'    oPres.Close()
'    oPres = Nothing
'    PwP.Quit()
'    PwP = Nothing
'    GC.Collect()
'End Sub

'Option Explicit On


'Sub MainDecoupage()
'    Formulaire.Show
'    If OnYVa = False Then Exit Sub
'    MsgBox "Attention ! Ce Proccessus peux etre long ! Merci de ne rien faire avec votre ordinateur avant le message de fin de traitement !", vbExclamation
'DossierExport = DossierExport & "\"

''Ouverture du fichier
'Set PPTdeBASE = Application.Presentations.Open(FileName:=FicherPPTdeBASE)

''boucle sur tout les fichiers du dossier
'Call ListFilesInFolder(DossierExport, True)

'    'on referme la slide de base
'    PPTdeBASE.Close

'    MsgBox "Voila, vos fichiers sont maintenant crées ! A bientôt !", vbExclamation
''Application.Quit
'End Sub
'Sub ajoutDeLaSlide(ByVal fichierPPTenCOUR As String)
'    Dim PPTenCOUR As PowerPoint.Presentation
'    Set PPTenCOUR = Application.Presentations.Open(FileName:=fichierPPTenCOUR)
'    PPTdeBASE.Slides(1).Copy
'    PPTenCOUR.Slides.Paste
'    PPTenCOUR.Save
'    PPTenCOUR.Close
'End Sub

'Sub test()
'    Call ListFilesInFolder("D:\OneDrive\Famille et amis\Balthazar", True)
'End Sub
'Sub ListFilesInFolder(ByVal xFolderName As String, ByVal xIsSubfolders As Boolean)
'    Dim xFileSystemObject As Object
'    Dim xFolder As Object
'    Dim xSubFolder As Object
'    Dim xFile As Object
'Set xFileSystemObject = CreateObject("Scripting.FileSystemObject")
'Set xFolder = xFileSystemObject.GetFolder(xFolderName)

'For Each xFile In xFolder.Files
'        Call ajoutDeLaSlide(xFile)
'    Next xFile
'    If xIsSubfolders Then
'        For Each xSubFolder In xFolder.SubFolders
'            Call ListFilesInFolder(xSubFolder.Path, True)
'        Next xSubFolder
'    End If
'Set xFile = Nothing
'Set xFolder = Nothing
'Set xFileSystemObject = Nothing
'End Sub

