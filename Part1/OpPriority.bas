Dim x, y As Double
Dim b As Boolean

Public Sub main()
x = 1: y = 2
'обчислення цього виразу виконується за правилом пріоритету операторів
b = (x ^ (2 + x) + 1) / (x * Cos(x + 1) - x) + x = 1 Or y > 0
Debug.Print b
'And має вищий пріоритет, Or - нижчий
Debug.Print True Or False And False ' True
Debug.Print (True Or False) And False ' False
End Sub
