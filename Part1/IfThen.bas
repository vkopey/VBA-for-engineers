Dim A, x, y As Double

Public Sub main()
x = 2: A = 1

'_____________1-й варіант конструкції_____________
If x > 0 Then y = x + A 'якщо x>0, то y=x+a

'_____________2-й варіант конструкції_____________
If x > 0 Then y = x + A Else y = x - A 'якщо x>0, то y=x+a, інакше y=x-a

'_____________3-й варіант конструкції_____________
If x > 0 Then 'якщо x>0, то
A = 2
y = x + A
Else 'інакше
A = 1
y = x - A
End If 'кінець умови

'_____________4-й варіант конструкції_____________
If x > 0 Then 'якщо x>0, то
A = 2
y = x + A
ElseIf x < 0 And x >= -5 Then 'інакше, якщо x<0 і x>=-5, то
A = 1
y = x - A
Else 'інакше
y = 0
End If 'кінець умови
End Sub