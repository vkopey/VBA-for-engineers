Type student 'тип користувача
    name As String * 20
    Ball As Double
End Type
Dim obj As student 'змінна типу student
Dim s1, s2 As String

'процедура main
Public Sub main()
'відкрити файл c:\file2.dat під номером 1 для довільного доступу; довжина запису файлу = Len(obj)
Open "c:\file2.dat" For Random As #1 Len = Len(obj)
Do 'початок циклу
    s1 = InputBox("Введіть ім'я", "Ім'я") 'ввести ім'я
    If s1 = "" Then Exit Do 'якщо нічого не введено, то вийти з циклу
    s2 = InputBox("Введіть бал", "Бал") 'ввести бал
    If s2 = "" Then Exit Do 'якщо нічого не введено, то вийти з циклу
    obj.name = s1 'присвоїти obj.Name значення
    obj.Ball = CDbl(s2) 'присвоїти obj.Ball значення
    Put #1, , obj 'записати obj у поточну позицію файлу 1
    Debug.Print obj.name; obj.Ball 'вивести
Loop 'повторити
Close #1 'закрити файл 1

'відкрити файл c:\file2.dat під номером 1 для довільного доступу
Open "c:\file2.dat" For Random As #1 Len = Len(obj)
Do While Not EOF(1) 'поки не кінець файлу 1
    Get #1, , obj 'прочитати дані з поточної позиції у obj
    'якщо знайдено ім'я "Ivanov" і не кінець файлу, то
    If Trim(obj.name) = "Ivanov" And Not EOF(1) Then _
        'вивести позицію запису і дані
        Debug.Print Seek(1) - 1; obj.name; obj.Ball
Loop 'повторити
Get #1, 1, obj 'прочитати дані з позиції 1
Debug.Print 1; obj.name; obj.Ball 'вивести дані
Close #1 'закрити файл
End Sub
