Public Sub main()
Debug.Print ActiveWorkbook.ActiveSheet.name 'ім'я активного листа активної книги
Debug.Print ActiveWorkbook.Worksheets(1).name 'ім'я першого листа активної книги
Debug.Print ActiveWorkbook.Saved 'чи збережено
Debug.Print Workbooks.Count 'кількість робочих книг у сімействі Workbooks
Workbooks.Add 'додати книгу
Workbooks(2).Activate 'активувати другу книгу
Workbooks(2).Password = "Пароль" 'установити пароль на другу книгу
Workbooks(2).Password = "" 'зняти пароль
Workbooks(2).SaveAs "my_book2" 'зберегти як my_book2.xls
Workbooks(2).Save 'зберегти
Debug.Print Workbooks(2).HasPassword 'чи має пароль
Workbooks(2).PrintOut 'вивести на друк
Workbooks(2).Close 'закрити
Workbooks.Open "my_book2" 'відкрити my_book2.xls
'змінити значення комірки A1 листа Лист1 книги my_book2
Workbooks("my_book2").Worksheets("Лист1").Range("A1").Value = 3
Workbooks("my_book2").Close 'закрити my_book2.xls
Kill "my_book2.xls" 'знищити файл my_book2.xls
End Sub
