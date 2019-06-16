Dim MatLab As Object 'об'єкт MatLab
Dim b(1, 2) As Double 'масив (реальна частина)
Dim z(1, 2) As Double 'нульовий масив (уявна частина)

Public Sub main()
Set MatLab = CreateObject("Matlab.Application") 'створити об'єкт MatLab
'заповнити масиви
For i = 0 To 1
    For j = 0 To 2
        b(i, j) = 1 'одиницями
        z(i, j) = 0 'нулями
    Next j
Next i
'виконати команди MatLab
MatLab.PutFullMatrix "b", "base", b, z 'передати матрицю 'b' в MatLab
MatLab.Execute "a = [1 2 3; 4 5 6]" 'створити матрицю 'a'
MatLab.Execute "b = a + b" 'додати матриці
'передати матрицю 'b' в VBA програму
MatLab.GetFullMatrix "b", "base", b, z
'вивести 'b'
Debug.Print b(0, 0), b(0, 1), b(0, 2)
Debug.Print b(1, 0), b(1, 1), b(1, 2)
MatLab.Quit 'вийти з MatLab
End Sub
