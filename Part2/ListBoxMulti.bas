Private Sub UserForm_Initialize()
Dim A(0 To 1, 0 To 1) As String 'масив
'заповнити масив
A(0, 0) = 1
A(0, 1) = "Перший"
A(1, 0) = 2
A(1, 1) = "Другий"
ListBox1.ColumnCount = 2 'кількість колонок
ListBox1.List = A 'заповнити список масивом
ListBox1.ColumnWidths = "20;20" 'ширина колонок
ListBox1.TextColumn = 2 'колонка, елемент якої повертається Text
End Sub

'процедура обробки події KeyDown (опущена клавіша на клавіатурі)
Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'якщо натиснуто "Insert"
If KeyCode = vbKeyInsert Then
    'якщо виключений мультивибір, то включити, і навпаки
    ListBox1.MultiSelect = IIf(ListBox1.MultiSelect = fmMultiSelectSingle, fmMultiSelectMulti, fmMultiSelectSingle)
End If
'якщо натиснуто "Delete"
If KeyCode = vbKeyDelete Then
    'i змінюється від 0 до кількості елементів-1
    For i = 0 To ListBox1.ListCount - 1
        'якщо елемент вибраний, додати його в ListBox2
        If ListBox1.Selected(i) Then ListBox2.AddItem ListBox1.List(i, 1)
    Next i
End If
End Sub
