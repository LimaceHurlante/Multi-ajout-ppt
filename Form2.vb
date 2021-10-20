Public Class Form2
    Public DossierExport As String = "", listeProvisoire
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
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
                CheckedListBox.Items.Add(Strings.Right(xFile.path, Strings.Len(xFile.path) - Strings.Len(DossierExport) - 1), True)
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
        DossierExport = FolderAddress()
        If DossierExport IsNot "" Then
            CheckedListBox.Items.Clear()
            LabelDirectoryName.Text = DossierExport & " :"
            Call ListFilesInFolder(DossierExport, True)
            If CheckedListBox.Items.Count > 0 Then
                ButtonSelectAll.Enabled = True
                ButtonUnselectAll.Enabled = True
                LabelNombreDeFichier.Text = "Nombre de fichiers selectionné(s) : " & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
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
        LabelNombreDeFichier.Text = "Nombre de fichiers selectionné(s) : " & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
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
        LabelNombreDeFichier.Text = "Nombre de fichiers selectionné(s) : " & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Form1.Fichiers.Add(CheckedListBox.SelectedItems)
        For i As Integer = 0 To CheckedListBox.CheckedItems.Count - 1
            Form1.Fichiers.Add(DossierExport & "\" & CheckedListBox.CheckedItems(i))

        Next i
        Me.Hide()
    End Sub

    Private Sub CheckedListBox_ChangeUICues(sender As Object, e As UICuesEventArgs) Handles CheckedListBox.ChangeUICues
        LabelNombreDeFichier.Text = "Nombre de fichiers selectionné(s) : " & CheckedListBox.CheckedIndices.Count & "/" & CheckedListBox.Items.Count
    End Sub
End Class