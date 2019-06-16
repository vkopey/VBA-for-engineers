'Option Explicit
DefStr S 'змінні, назва яких починається з S, мають тип string
'описати змінні з типом:
Dim i1 As Byte 'байт (коротке ціле від 0 до 255, розміром 1 байт)
Dim b As Boolean 'логічний (булевий) (значення: true (або 1), false (або 0))
Dim i2 As Integer 'цілий (ціле в межах +-32768, розміром 2 байти)
Dim i3 As Long 'довгий цілий (розміром 4 байти)
Dim x1 As Single 'дійсний звичайної точності (розміром 4 байти)
Dim x, y As Double 'дійсний подвійної точності (розміром 8 байт)
Dim d As Date 'календарна дата (розміром 8 байт)
Dim obj As Object 'об'єкт (розміром 4 байти)
Dim obj2 As New Worksheet 'об'єкт робочий лист Excel
Dim s As String 'рядок
Dim s2 As String * 10 'рядок розміром 10 символів
Dim x2 As Variant 'числові підтипи (розміром 16 байт)
Private Type student 'тип користувача, який описує поняття студента -
    number As Integer 'його номер залікової книжки
    name As String 'і ім'я
End Type 'кінець опису типу
Dim obj3 As student 'описати змінну obj3 з типом student
Const s3 = "Hello!" 'константа
Public Const pi As Double = 3.14 'константа, видима в усіх модулях

Public Sub main() 'підпрограма-процедура з іменем main,
'видима в усіх модулях проекту
'присвоїти змінним значення
i1 = 1 'ціле
b = True 'логічне
i2 = 12500 'ціле
i3 = 256132 'ціле
x1 = 5.124 'дійсне
x = 34.345 'дійсне
y = -25.684 'дійсне
d = Date 'присвоїти поточну дату, наприклад 21.09.2008
'присвоїти об'єкту obj вказівник на активну комірку Excel
Set obj = Excel.ActiveCell
obj.Value = 1 'присвоїти властивості Value значення 1
s = "hello world!" 'рядок
f$ = "Програмування на VBA" 'рядок
x2 = 54.76 'дійсне
x3 = 398 'ціле (змінна x3 не описана, тому її тип Variant). Змінна буде не визначена, якщо забрати примітку з Option Explicit на початку модуля.
obj3.number = 1 'полю number ціле
obj3.name = "Іванов" 'полю name рядок
'вивести значення даних у вікно Immediate
Debug.Print i1; b; i2; i3; x1; x; y; d; obj.Value; s; f$; x2; x3; obj3.name
End Sub 'кінець підпрограми main
