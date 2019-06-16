Private Sub UserForm_Initialize()
ScrollBar1.Orientation = fmOrientationHorizontal 'орієнтація
ScrollBar1.Min = 0 'мінімальне значення
ScrollBar1.Max = 10 'максимальне значення
ScrollBar1.SmallChange = 1 'малий крок зміни значення
ScrollBar1.LargeChange = 2 'великий крок зміни значення
ScrollBar1.Delay = 10 'затримка подій зміни значення
SpinButton1.Min = 0 'мінімальне значення
SpinButton1.Max = 10 'максимальне значення
SpinButton1.SmallChange = 1 'малий крок зміни значення
End Sub

'процедура обробки події Change (зміна значення)
Private Sub ScrollBar1_Change()
'змінити надпис на формі на значення ScrollBar1
UserForm11.Caption = ScrollBar1.Value
End Sub

'процедура обробки події SpinUp (натиснуто кнопку "вверх")
Private Sub SpinButton1_SpinUp()
'збільшити значення на 2
SpinButton1.Value = SpinButton1.Value + 2
End Sub

'процедура обробки події Change (зміна значення)
Private Sub SpinButton1_Change()
'змінити надпис на формі на значення SpinButton1
UserForm11.Caption = SpinButton1.Value
End Sub
