Public Sub main()
ActiveSheet.UsedRange.Clear 'очистити діапазон з значеннями
ActiveSheet.UsedRange.ClearContents 'очистити вміст
ActiveSheet.UsedRange.ClearComments 'очистити коментарі
ActiveSheet.UsedRange.ClearFormats 'очистити формати
'змінити значення комірки A1 листа Лист1 книги VBA_examples
Workbooks("VBA_examples").Worksheets("Лист1").Range("A1").Value = 5
Range("Лист1!A1").Value = 5 'або так
Range("A1").Value = 5 'або так
Cells(1, 1).Value = 5 'або так
Rows(10).RowHeight = StandardHeight 'стандартна висота 10-го рядка
Columns(10).ColumnWidth = StandardWidth 'стандартна ширина 10-го стовпця
Range("E5").EntireRow.Clear 'очистити рядок 5
Range("E5").EntireColumn.Clear 'очистити стовпчик E
Range("A1").AddComment "Коментар" 'додати коментар
Debug.Print Range("A1").Comment.Text 'текст коментаря
Range("A1").Font.Size = 12 'розмір шрифту
Range("A1").Font.Color = RGB(255, 0, 0) 'колір шрифту
Range("A1:B2").name = "MyRange" 'назва діапазону
Debug.Print Range("A1:B2").Count 'кількість комірок
Range("B1:B2").Formula = "=$A$1+1" 'формула
Range("B1:B2").FormulaR1C1 = "=(R1C1)+1" 'формула в форматі R1C1
Range("C1:D2").FormulaArray = "=TRANSPOSE(A1:B2)" 'формула масиву
Range("E1").FormulaLocal = "=СУММ(C1:D2)" 'формула неангломовної версії Excel
Debug.Print Range("E2").Text 'вміст у текстовому форматі
Debug.Print Range("A1").Address(True, False) 'адреса
Debug.Print Range("C1:D2").Rows.Count 'кількість рядків
Debug.Print Range("C1:D2").Columns.Count 'кількість стовпців
Range("A1").EntireColumn.AutoFit 'авторозмір
Range("A3:A4").Cut 'вирізати в буфер обміну
Range("A1").Copy 'скопіювати в буфер обміну
Range("A4").PasteSpecial xlPasteValues 'вставити з буфера значення
Range("H10").Delete 'знищити
Rows(1).Insert 'вставити новий рядок перед першим рядком
Range("A1").Offset(2, 0).Value = 1 'змінити значення комірки A3
Range("A4:A5").Select 'виділити діапазон
Selection.Copy 'копіювати виділення в буфер обміну

'методи для роботи з діапазонами
'заголовки полів
Range("F1").Value = "Значення": Range("G1").Value = "Номер"
Range("F2").Value = 1 'перше значення
'геометрична прогресія xlGrowth з кроком 2 і кінцевим значенням 16
Range("F2").DataSeries xlColumns, xlGrowth, Step:=2, Stop:=16
'перші значення
Range("G2").Value = 1
Range("G3").Value = 2
Range("G2:G3").AutoFill Range("G2:G6") 'автозаповнення: 1,2,3,4,5
'автофільтр по другому стовпчику, значення 2 або 3
Range("F2:G6").AutoFilter 2, "2", xlOr, "3"
Range("F2:G6").AutoFilter 'відмінити автофільтр
'заголовки полів для критеріїв
Range("F7").Value = "Значення": Range("G7").Value = "Номер"
Range("G8").Value = 4 'критерій
'розширений фільтр копіює знайдене (за критерієм F7:G8) в діапазон H1:I1
Range("F1:G6").AdvancedFilter xlFilterCopy, Range("F7:G8"), Range("H1:I1")
Range("F2:G6").Sort Range("G2"), xlDescending 'сортування за спаданням
Range("G2:G6").Find("4").Activate 'активувати комірку зі знайденим значенням 4
End Sub
