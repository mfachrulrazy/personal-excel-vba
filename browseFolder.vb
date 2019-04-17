Sub browseFolderPath()
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFolderPicker)

    'Mencegah Memilih lebih dari 1 folder
    fileExplorer.AllowMultiSelect = False

    With fileExplorer
        If .Show = -1 Then 'Pilih Folder
            AbsLokasi = .SelectedItems.Item(1)
            Path = AbsLokasi & "\"
        Else
            MsgBox "Pilih Folder dibatalkan"
            AbsLokasi = "" 'Ketika dibatalkan/cancel
        End If
    End With
End Sub