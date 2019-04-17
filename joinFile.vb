Sub GabungFile()

'Mengambil semua nama file
Filename = Dir(AbsLokasi & "*.xlsx")
'Deklarasi Variable
Dim ws As Worksheet
Dim wbnew As Workbook

Set wbnew = Workbooks.Add

'Perulangan do jika Filename tidak kosong
Do While Filename <> ""
    Workbooks.Open Filename:=AbsLokasi & Filename, ReadOnly:=True
    For Each Sheet In ActiveWorkbook.Sheets
        Sheet.Copy After:=wbnew.Sheets(1)
    Next Sheet
    Workbooks(Filename).Close
    Filename = Dir()
Loop
ActiveWorkbook.SaveAs AbsLokasi & "FileBaru.xlsx"
End Sub