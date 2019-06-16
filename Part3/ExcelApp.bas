Public Sub main()
'деякі властивості Application:
Debug.Print Application.ActiveWorkbook.Name 'ім'я активної книги, або скорочено ActiveWorkbook.Name
Debug.Print Application.ActiveSheet.Name 'ім'я активного листа
Debug.Print Application.ActiveCell.Value 'значення активної комірки
'ім'я активного листа книги, в якій виконується макрос
Debug.Print ThisWorkbook.ActiveSheet.Name
Application.Calculation = xlCalculationAutomatic 'режим обчислень
Application.Caption = "Моя програма" 'надпис
Application.Caption = Empty 'надпис за замовчуванням
Application.Cells(1, 1).Value = 1 'значення комірки Excel (1,1)
Application.DisplayStatusBar = True 'показувати рядок стану
Application.StatusBar = "Hello!" 'текст у рядку стану
Application.ScreenUpdating = False 'не оновлювати екран
Application.ScreenUpdating = True 'оновлювати екран
Debug.Print Application.Version 'версія Excel
Application.EnableCancelKey = xlInterrupt 'переривати виконання натиском Ctrl-Break
Application.WindowState = xlMaximized 'стан вікна

'деякі методи Application:
Application.Calculate 'обчислити книгу
'вивести діалогове вікно, результат присвоїти і
i = InputBox("Виконати (1),відкласти виконання (2),вийти (3)", "Вибір", 0)
Select Case i 'вибір і
Case 1
Application.Run "VBAProject.Module1.main" 'виконати макрос
Case 2
'відкласти виконання макроса на 10 секунд
Application.OnTime Now + TimeValue("0:00:10"), "Module1.main"
Case 3
Application.Quit 'вийти з Excel
Case Else
End Select
Application.OnKey "^{a}", "Module1.main" 'виконати макрос піля натиску Ctrl-A
'конвертувати формулу з формату R1C1 у формат A1
Debug.Print Application.ConvertFormula("=SUM(R1C1:R5C1)", xlR1C1, xlA1)
'виділити перетин діапазонів
Application.Intersect(Range("A1:B2"), Range("B2:C3")).Select
'виділити об'єднання діапазонів
Application.Union(Range("A1:B2"), Range("B2:C3")).Select
answ = Application.Dialogs(xlDialogOpen).Show 'показати діалогове вікно відкриття файлу

'Evaluate конвертує ім'я Excel в об'єкт або значення
Application.Evaluate("A2").Value = 2 'перетворити рядок в об'єкт
[A2].Value = 2 ' або
Debug.Print Application.Evaluate("SUM(A1:A2)") ' перетворити рядок в функцію
'Debug.Print Application.Evaluate("SUM(1,2)") ' або
Set r = Range("A1:A2")
Debug.Print Application.Sum(r) ' виклик вбудованої функції листа Excel
'Debug.Print Application.WorksheetFunction.Sum(r) ' або
'Debug.Print Application.Sum(1, 2) ' або

'Debug.Print my_funct(3) 'виклик функції користувача листа Excel
End Sub
