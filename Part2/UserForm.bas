'Процедура обробки події Initialize (ініціалізація форми)
Private Sub UserForm_Initialize()
'приклади властивостей UserForm:
UserForm1.Caption = "Перша форма" 'надпис
UserForm1.BackColor = vbGreen 'колір фону
UserForm1.BorderStyle = fmBorderStyleSingle 'стиль границі
UserForm1.Picture = LoadPicture("C:\WINDOWS\Паркет.bmp") 'фонова картинка
UserForm1.PictureSizeMode = fmPictureSizeModeStretch 'розмір картинки
UserForm1.StartUpPosition = 0 'початкова позиція
'координати верхнього лівого кута:
UserForm1.Left = 50
UserForm1.Top = 50
UserForm1.Height = 100 'висота
UserForm1.Width = 200 'ширина
UserForm1.MousePointer = fmMousePointerCross 'вид вказівника миші
UserForm1.Enabled = True 'чи допустиме керування вручну
'UserForm1.ShowModal = False 'зробити не модальною. Властивість змінюється тільки на етапі проектування!
End Sub

'Процедура обробки події Click (натиск лівої кнопки миші)
Private Sub UserForm_Click()
'приклади методів UserForm:
UserForm1.Hide 'сховати
MsgBox ("UserForm1.Show") 'вивести вікно з повідомленням
UserForm1.Move 0, 0, 300, 400 'перемістити і змінити розмір
UserForm1.Show 'показати
End Sub

'Процедура обробки події Terminate (знищення)
Private Sub UserForm_Terminate()
MsgBox ("Відбулась подія Terminate") 'вивести вікно з повідомленням
End Sub
