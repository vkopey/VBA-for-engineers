Dim x, y As Double

Public Sub main()
x = 2.7
Select Case x 'вибір стосується змінної 'x'
Case 1.5 'якщо x=1.5, то
y = 7.4
Case 2 To 2.7, 3, Is > 4 'якщо 2<=x<=2.7 або x=3 або x>4, то
y = 3.2
Case Else 'у інших випдках
y = 0
End Select 'кінець вибору
Debug.Print y
End Sub
