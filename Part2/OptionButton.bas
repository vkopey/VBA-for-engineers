Private Sub UserForm_Initialize()
OptionButton3.Value = True 'значення
OptionButton3.Caption = "True" 'надпис
End Sub

'Процедура обробки події Click (натиск лівої кнопки миші)
Private Sub CommandButton1_Click()
'надпис рамки змінити на ім'я активного елемента
Frame1.Caption = Frame1.ActiveControl.name
For Each opt In Frame1.Controls 'для кожного елемента
    opt.Caption = opt.Value 'змінити надпис на його значення
Next opt
End Sub
