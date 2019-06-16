Private Sub UserForm_Initialize()
CheckBox1.Enabled = True 'доступність
CheckBox1.TripleState = True 'дозволити три стани
CheckBox1.Value = True 'значення "вибрано"
CheckBox1.Value = False 'значення "не вибрано"
CheckBox1.Value = Null 'значення "третій стан"
CheckBox1.SetFocus 'установити фокус
ToggleButton1.Value = False 'значення "не вибрано"
End Sub

'Процедура обробки події Change (зміна стану)
Private Sub CheckBox1_Change()
If CheckBox1.Value Then 'якщо значення=True
    CheckBox1.Caption = "True" 'змінити надпис на "True"
ElseIf Not CheckBox1.Value Then 'інакше, якщо значення=False
    CheckBox1.Caption = "False" 'змінити надпис на "False"
Else: CheckBox1.Caption = "Null" 'інакше змінити надпис на "Null"
End If
End Sub
