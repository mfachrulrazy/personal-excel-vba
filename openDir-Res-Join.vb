Public Path

Sub browseFolderPath()
    Dim fileExplorer As FileDialog
    Set fileExplorer = Application.FileDialog(msoFileDialogFolderPicker)

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


Sub GabungFile()
AbsLokasi = Path
Filename = Dir(AbsLokasi & "*.xlsx")
Dim ws As Worksheet
Dim wbnew As Workbook
Do While Filename <> ""
    Workbooks.Open Filename:=AbsLokasi & Filename, ReadOnly:=True
    For Each Sheet In ActiveWorkbook.Sheets
        Sheet.Copy After:=ThisWorkbook.Sheets(1)
    Next Sheet
    Workbooks(Filename).Close
    Filename = Dir()
Loop
ThisWorkbook.Worksheets("Sheet1").Select
End Sub

Public wks As Worksheet
Sub DelCol()

For Each wks In Application.ThisWorkbook.Worksheets
   If wks.Name <> "Sheet1" Then
    wks.Select
    Range("A:A,C:C,D:D,G:G,H:H,I:I,L:L,M:M").Select
    Range("M1").Activate
    Selection.Delete Shift:=xlToLeft
   End If
Next
ThisWorkbook.Worksheets("Sheet1").Select
End Sub

Sub delSht()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each wks In Application.ThisWorkbook.Worksheets
    If wks.Name <> "Sheet1" Then
        wks.Delete
    End If
Next
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub