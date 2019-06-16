Dim x As Double

Public Sub main()
x = 0
Debug.Assert x <> 0 'якщо 'x' дорівнює 0, призупинити виконання
On Error GoTo HandleError 'при помилці перейти на мітку HandleError
Debug.Print 1 / x 'помилка 11 (ділення на нуль)
x = 1E+300 * 1E+300 'помилка 6 (переповнення)
x = CDbl("0.12") 'помилка 13 (невідповідність типу)
Err.Raise 65535 'створити помилку виконання 65535
Exit Sub 'вийти з підпрограми
HandleError: 'мітка
Select Case Err.number 'якщо номер помилки
    Case 11 'рівний 11 (ділення на нуль)
        x = 1 'змінити знаменник
        Debug.Print "Ділення на нуль!" 'вивести повідомлення
        Resume 'повторити інструкцію з помилкою
    Case Else 'інший номер
        Debug.Print Err.number 'вивести номер помилки
        Debug.Print Err.Description 'вивести опис помилки
End Select
Resume Next 'перейти на наступну інструкцію за помилкою
End Sub
