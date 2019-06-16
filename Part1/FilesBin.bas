Public Sub main()
'відкрити файл c:\file1.dat під номером 1 як бінарний
Open "c:\file1.dat" For Binary As #1
Do While Not EOF(1) 'поки не кінець файлу 1
    x = Input(1, #1) 'присвоїти 'x' наступний байт з файлу 1
    Debug.Print x; 'вивести 'x'
Loop 'повторити
Close #1 'закрити файл 1
End Sub
