Public Class Form2
    Public DossierDeTravail As String = "", dossierSauvegarde As String, checklistSauvegardeItemName As ArrayList, checklistSauvegardeItemChecked As ArrayList
    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        Me.Hide()
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
                'Debug.Print(xFile.name)
                'Form1.Fichiers.Add(xFile.path)
                CheckedListBox.Items.Add(Strings.Right(xFile.path, Strings.Len(xFile.path) - Strings.Len(DossierDeTravail) - 1), True)
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

    Private Sub ButtonPickFolder_Click(sender As Object, e As EventArgs) Handles ButtonPickFolder.Click
        DossierDeTravail = FolderAddress()
        Call ExploreLeDossier()


    End Sub
    Sub ExploreLeDossier()

        If DossierDeTravail IsNot "" Then
            CheckedListBox.Items.Clear()
            LabelDirectoryName.Text = DossierDeTravail & " :"
            Call ListFilesInFolder(DossierDeTravail, IncludeSubFolder.Checked)
            If CheckedListBox.Items.Count > 0 Then
                ButtonSelectAll.Enabled = True
                ButtonUnselectAll.Enabled = True
                LabelNombreDeFichier.Text = Form1.FileSelected & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
            End If
        Else
            MsgBox(Form1.NoSelectedFolder)
        End If

        'Call CheckIfReadyToGo()
    End Sub



    Private Function FolderAddress() As String
        Dim dlg As New FolderBrowserDialog With {
            .Description = Form1.ChooseFolder
        }
        If dlg.ShowDialog() = DialogResult.OK Then
            Dim path As String = dlg.SelectedPath
            Return path
        Else
            Return String.Empty
        End If
    End Function
    Private Sub CheckedListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox.SelectedIndexChanged
        LabelNombreDeFichier.Text = Form1.FileSelected & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
    End Sub

    Private Sub ButtonSelectAll_Click(sender As Object, e As EventArgs) Handles ButtonSelectAll.Click
        For i As Integer = 0 To CheckedListBox.Items.Count - 1
            CheckedListBox.SetItemChecked(i, True)
        Next i
    End Sub

    Private Sub ButtonUnselectAll_Click(sender As Object, e As EventArgs) Handles ButtonUnselectAll.Click
        For i As Integer = 0 To CheckedListBox.Items.Count - 1
            CheckedListBox.SetItemChecked(i, False)
        Next i
    End Sub

    Private Sub CheckedListBox_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox.ItemCheck
        LabelNombreDeFichier.Text = Form1.FileSelected & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
    End Sub

    Private Sub IncludeSubFolder_CheckedChanged(sender As Object, e As EventArgs) Handles IncludeSubFolder.CheckedChanged
        If DossierDeTravail IsNot "" Then
            Call ExploreLeDossier()
        End If
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LabelDirectoryName.Text = ""
    End Sub

    Private Sub ButtonValider_Click(sender As Object, e As EventArgs) Handles ButtonValider.Click
        If CheckedListBox.CheckedItems.Count > 0 Then Form1.Fichiers.Clear()
        dossierSauvegarde = DossierDeTravail

        For i As Integer = 0 To CheckedListBox.CheckedItems.Count - 1
            Form1.Fichiers.Add(DossierDeTravail & "\" & CheckedListBox.CheckedItems(i))
        Next i
        Me.Hide()
    End Sub

    Private Sub CheckedListBox_ChangeUICues(sender As Object, e As UICuesEventArgs) Handles CheckedListBox.ChangeUICues
        LabelNombreDeFichier.Text = Form1.FileSelected & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
    End Sub
    Private Sub Form2_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If Not IsNothing(dossierSauvegarde) Then
            DossierDeTravail = dossierSauvegarde
            LabelDirectoryName.Text = DossierDeTravail & " :"
            'CheckedListBox.Item = checklistSauvegarde
        End If
    End Sub
End Class