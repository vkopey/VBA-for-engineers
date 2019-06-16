Dim y As Integer

'Головна підпрограма-процедура main
Public Sub main()
y = Sum(2, 3) 'присвоїти 'y' значення функції Sum з параметрами 2, 3
Debug.Print y 'вивести 'y'
Debug.Print Sum2() 'результат: 1
Debug.Print Sum2() 'результат: 2
Debug.Print Sum2() 'результат: 3
Debug.Print Fact(3) 'вивести факторіал 3
End Sub 'кінець процедури

'Функція Sum, яка повертає значення цілого типу
Public Function Sum(A, b As Integer) As Integer 'a, b - параметри
Sum = A + b 'повернути суму a + b
End Function 'кінець функції

'Функція з статичною локальною змінною
Public Function Sum2() As Integer
Static n As Integer 'статична змінна зберігає своє значення
n = n + 1 'змінити значення статичної змінної
Sum2 = n 'повернути N
End Function

'Рекурсивна функція (викликає сама себе) для обчислення факторіала
Public Function Fact(n As Integer)
'якщо N < 1, то Fact = 1, інакше викликати Fact(N - 1) і помножити на N
If n < 1 Then Fact = 1 Else Fact = Fact(n - 1) * n
End Function
