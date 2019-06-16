Public Sub main()
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Set xlApp = CreateObject("Excel.Application") ' Excel
xlApp.Workbooks.Add ' нова робоча книга
Set xlBook = xlApp.ActiveWorkbook ' активна книга
'Set xlBook = xlApp.Workbooks.Open("e:\mytestbook.xls") ' або відкрити
Set xlSheet = xlBook.Worksheets(1) ' перший лист
xlSheet.Cells(1, 1).Value = 1 ' записати в комірку
'xlApp.Visible = True ' зробити видимим
xlBook.SaveAs ("e:\mytestbook.xls") 'зберегти як
'xlBook.Save ' або зберегти
xlApp.Quit ' вийти з Excel
End Sub
