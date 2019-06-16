Private Sub UserForm_Initialize()
TabStrip1.MultiRow = True 'дозволити кілька рядків вкладок
TabStrip1.TabOrientation = fmTabOrientationLeft 'орієнтація
TabStrip1.TabIndex = 1 'вибрано вкладку з індексом 1
TabStrip1.Tabs.Item(0).Caption = "A" 'надпис першої вкладки
TabStrip1.Tabs.Item(1).Caption = "B" 'надпис другої вкладки
TabStrip1.Tabs.Add "Tab3", "C", 2 'додати третю вкладку
End Sub

'процедура обробки події Change (зміна вкладки)
Private Sub TabStrip1_Change()
'змінити надпис на індекс вибраної вкладки
Label1.Caption = TabStrip1.SelectedItem.Index
End Sub
