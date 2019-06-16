Dim x, y As Integer 'глобальні змінні
Dim B1(0 To 2), B2(2), B3(2) As Integer 'глобальні масиви

'Головна підпрограма-процедура main
Public Sub main()
'Приклад 1
Sum 2, 3, y 'виклик процедури Sum з параметрами 2, 3, y
'або Call Sum(2, 3, y)
'або Sum A:=2, B:=3, C:=y
Debug.Print y 'виведення 'y'

'Приклад 2
x = 1: y = 1
Sum2 x, y 'виклик процедури Sum2. Результат: x=3, y=1
Debug.Print x; y 'виведення

'Приклад 3
B1(0) = 1: B1(1) = 5: B1(2) = 3 'заповнити масив B1
B2(0) = 9: B2(1) = 5: B2(2) = 7 'заповнити масив B2
Sum3 B1, B2, B3 'виклик процедури Sum3
Debug.Print B3(0), B3(1), B3(2) 'виведення

'Приклад 4
Sum4 2, 3 'виклик процедури Sum4

'Приклад 5
Sum5 1, x, y 'виклик процедури Sum5
Debug.Print x, y 'виведення
End Sub 'кінець процедури

'Процедура Sum
Public Sub Sum(A, b, c As Integer) 'a, b, c - параметри
c = A + b 'тіло процедури
End Sub 'кінець процедури

'Процедура Sum2
'Параметр A передається за посиланням (за замовчуванням), B - за значенням
Public Sub Sum2(ByRef A As Variant, ByVal b As Integer)
Dim n As Integer 'локальна змінна
n = 2
A = A + n 'A - синонім 'x'
b = b + n 'B - окрема копія 'y'
End Sub

'Процедура Sum3
Public Sub Sum3(A1(), A2(), A3() As Integer) 'параметри - масиви
For i = 0 To 2
    A3(i) = A1(i) + A2(i) 'додати масиви A1 і A2
Next i
End Sub

'Процедура Sum4
'Параметр C не обов'язковий, за замовчуванням рівний 1
Public Sub Sum4(A, b As Integer, Optional c As Integer = 1)
'якщо не вказано C і A=0, вийти з процедури
If IsMissing(c) And A = 0 Then Exit Sub
Debug.Print A + b + c 'вивести суму
End Sub

'Процедура Sum5
Public Sub Sum5(A As Integer, ParamArray z()) 'необмежена кількість параметрів
For i = LBound(z) To UBound(z)
    z(i) = z(i) + A 'додати до кожного параметра A
Next i
End Sub
