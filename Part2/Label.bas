'Процедура обробки події Click (натиск лівої кнопки миші)
Private Sub Label1_Click()
Label1.AutoSize = False 'авторозмір
Label1.Height = 40 'висота
Label1.Width = 70 'ширина
Label1.Font.name = "Times New Roman" 'ім'я шрифту
Label1.Font.Size = 14 'розмір шрифту
Label1.Font.Bold = True 'жирний шрифт
Label1.Font.Italic = True 'курсив шрифт
Label1.Font.Underline = True 'підкреслений шрифт
Label1.ForeColor = RGB(255, 0, 0) 'колір переднього плану (шрифту)
Label1.TextAlign = fmTextAlignCenter 'вирівнювання надпису
Label1.SpecialEffect = fmSpecialEffectSunken 'спеціальний візуальний ефект
Label1.Caption = "Clicked" 'надпис
Label1.WordWrap = True 'перенос тексту
End Sub
