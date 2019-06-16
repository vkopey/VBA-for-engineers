Private Sub UserForm_Initialize()
ListBox1.ListStyle = fmListStyleOption 'стиль списку
ListBox1.TextAlign = fmTextAlignCenter 'вирівнювання тексту
ListBox1.MatchEntry = fmMatchEntryFirstLetter 'пошук по першій букві
'перший спосіб заповнення списку:
ListBox1.AddItem "Перший" 'перший елемент списку
ListBox1.AddItem "Другий" 'другий елемент списку
ListBox1.AddItem "Третій" 'третій елемент списку
ListBox1.Clear 'очистити список
'другий спосіб заповнення списку:
ListBox1.List = Array("Перший", "Другий", "Третій")
End Sub

'процедура обробки події Click (натиск лівою кнопкою миші)
Private Sub ListBox1_Click()
'вивести в надпис форми вибраний елемент, його індекс, кількість елементів
UserForm8.Caption = ListBox1.Text & " " & ListBox1.ListIndex _
& "/" & ListBox1.ListCount
End Sub

'процедура обробки події KeyDown (опущена клавіша на клавіатурі)
Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'якщо натиснуто "Insert", додати новий елемент
If KeyCode = vbKeyInsert Then ListBox1.AddItem "Новий", ListBox1.ListIndex + 1
'якщо натиснуто "Delete", видалити елемент
If KeyCode = vbKeyDelete Then ListBox1.RemoveItem ListBox1.ListIndex
End Sub
