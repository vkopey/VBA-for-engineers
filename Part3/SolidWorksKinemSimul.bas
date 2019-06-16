Const swDocPART = 1 'константи \SldWorks\samples\appComm\swconst.h
Const swDocASSEMBLY = 2
Const swDocDRAWING = 3
Dim swApp As Object 'оголошення об'єкта swApp
Dim Part As Object 'оголошення об'єкта Part
Dim x As Double 'оголошення змінної параметра моделі
Dim i As Integer 'оголошення змінної лічильника записів бази даних
'Ініціалізується форма користувача
Private Sub UserForm_Initialize()
    MyPath = CurDir 'визначення поточного каталогу
    MyPath = "C:\sldworks\ПСВ"
    'створення об’єкта swApp
    Set swApp = CreateObject("SldWorks.Application")
    'створення об’єкта Part
    Set Part = swApp.OpenDoc(MyPath + "\ПСВ.SLDASM", swDocASSEMBLY)
    If Part Is Nothing Then
       Exit Sub
       'якщо Part не існує, то вихід
    Else
      'інакше активувати
       Set Part = swApp.ActivateDoc("ПСВ.SLDASM")
    End If
    'початкові параметри елементів діалогового вікна
    TextBox1.Text = "D1@Угол2"
    TextBox5.Text = "RD1@Примечания"
    OptionButton1.Value = True
    TextBox2.Text = "0"
    TextBox3.Text = "1"
    TextBox4.Text = "0,1"
End Sub 
'Натиснута кнопка CommandButton1
Private Sub CommandButton1_Click()
'якщо включений перемикач OptionButton1, то p_form
 If OptionButton1.Value = True Then p_form
'якщо включений перемикач OptionButton2, то p_tabl
 If OptionButton2.Value = True Then p_tabl
End Sub
'Процедура виконує табулювання кінематичних параметрів для заданих у формі параметрів переміщень
Private Sub p_form()
Dim p1, p2 As String
    Dim xp, xk, xd As Double
    i = 3 'початкове значення лічильника
    'отримати дані з текстових полів
    p1 = TextBox1.Text
    p2 = TextBox5.Text
    xp = CDbl(TextBox2.Text)
    xk = CDbl(TextBox3.Text)
    xd = CDbl(TextBox4.Text)
    'очистити діапазон
    Range("b3:b1000").Clear
    Range("e3:e1000").Clear
    'зміна значень параметрів
    For x = xp To xk Step xd
    Part.Parameter(p1).SystemValue = x
    Part.EditRebuild 'Перебудова моделі
    'Виведення значень параметрів у відповідні комірки Excel
    Cells(i, 2).Value = x
    Cells(i, 5).Value = Part.Parameter(p2).SystemValue * 1000
    i = i + 1 'зміна значення лічильника
    Next x
End Sub
'Процедура виконує табулювання кінематичних параметрів для заданих у таблиці параметрів переміщень, швидкостей і прискорень
Private Sub p_tabl()
   Dim p1, p2 As String
    'отримати текст з текстових полів
    p1 = TextBox1.Text
    p2 = TextBox5.Text
    'очистити діапазон
    Range("e3:e1000").Clear
    'цикл від трьох до кількості чисел в діапазоні В
    For i = 3 To Cells(1, 8).Value
    'зміна значень параметрів
    Part.Parameter(p1).SystemValue = Cells(i, 2).Value
    Part.EditRebuild 'перебудова моделі
    'Виведення значень параметрів у відповідні комірки Excel
    Cells(i, 5).Value=Part.Parameter(p2).SystemValue*1000
    Next i
End Sub
'Процедура закриття діалогового вікна
Private Sub CommandButton2_Click()
UserForm1.Hide
End Sub
