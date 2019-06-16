Dim x, y As Double
Dim s As String

'підпрограма main
Public Sub main()
x = 5.6
'відкрити файл c:\file1.dat під номером 1 для виведення
Open "c:\file1.dat" For Output As #1
Print #1, "значення x="; x 'записати форматовані дані у файл 1
Write #1, x ^ 2 'записати неформатовані дані у файл 1
Close #1 'закрити файл 1

'відкрити файл c:\file1.dat для введення під номером 1
Open "c:\file1.dat" For Input As #1 '
Line Input #1, s 'прочитати рядок з файлу
Input #1, y 'прочитати дане з файлу
Debug.Print s; y 'вивести s,y у вікно Immediate
Close #1 'закрити файл 1

'відкрити файл c:\file1.dat під номером 1 для додання
Open "c:\file1.dat" For Append As #1
Write #1, x ^ 3 'записати неформатовані дані у файл 1
Close #1 'закрити файл 1
End Sub
