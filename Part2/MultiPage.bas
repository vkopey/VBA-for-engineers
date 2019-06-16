Dim TextBox1 As Control 'об'єкт TextBox1

Private Sub UserForm_Initialize()
MultiPage1.Pages(0).Caption = "A" 'змінити надпис сторінки
MultiPage1.Pages(1).Caption = "B" 'змінити надпис сторінки
'створити новий об'єкт TextBox1
Set TextBox1 = MultiPage1.Pages(1).Controls.Add("Forms.TextBox.1", "TextBox1", Visible)
TextBox1.Visible = True 'зробити видимим
End Sub

'процедура обробки події Click
Private Sub CommandButton1_Click()
TextBox1.Text = "Hello!" 'змінити текст
End Sub

'процедура обробки події Change (сторінка змінена)
Private Sub MultiPage1_Change()
'змінити надпис форми на індекс вибраної сторінки
UserForm13.Caption = MultiPage1.SelectedItem.Index
End Sub
