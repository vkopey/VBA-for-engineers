Dim i, s As Integer

Public Sub main()
'знайти суму цілих чисел від 1 до 100
s = 0 'сума = 0
i = 1 'перше число
While i <= 100 'поки і менше рівне 100
s = s + i 'додати до суми 'i'
i = i + 1 'наступне 'i'
Wend 'повторити
Debug.Print s 'вивести суму
End Sub
