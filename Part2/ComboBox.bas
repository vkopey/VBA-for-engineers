Private Sub UserForm_Initialize()
ComboBox1.List = Array(1, 2, 3, 4, 5, 6, 7) 'заповнити список
ComboBox1.ListRows = 4 'у списку показувати 4 рядка
ComboBox1.MatchRequired = True 'заборона введення у текстове поле значень, яких немає у списку
End Sub

'процедура обробки події Change (зміна значення текстового поля)
Private Sub ComboBox1_Change()
'вивести у надпис форми текст текстового поля і індекс вибраного елемента
UserForm10.Caption = ComboBox1.Text & " " & ComboBox1.ListIndex
End Sub
