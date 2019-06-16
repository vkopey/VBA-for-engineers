'Оголосити об'єкт класу Class1
Dim WithEvents obj As Class1 ' об'єкт може викликати події

Private Sub CommandButton1_Click()
Set obj = New Class1 'створити об'єкт obj
obj.z = 1
obj.x = TextBox1.Text 'тут може виникнути подія notNumber
TextBox1.Text = obj.y(5) 'вивести результат в TextBox1
End Sub

'Обробник події користувача notNumber
Private Sub obj_notNumber(x As Variant)
MsgBox "Warning! x is not number: " & x
End Sub
