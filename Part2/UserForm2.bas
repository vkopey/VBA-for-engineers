Private Sub UserForm_Initialize()
'для кожного елемента керування на UserForm2
For Each obj In UserForm2.Controls
    obj.AutoSize = True 'авторозмір
    obj.Visible = True 'видимість
    obj.Enabled = True 'дозвіл керування
    'координати верхнього лівого кута
    obj.Left = 10
    obj.Top = obj.Top + 20
    obj.Height = 20 'висота
    obj.Width = 100 'ширина
    obj.ControlTipText = "help" 'текст підказки
    obj.BackColor = vbYellow 'колір фону
    obj.ForeColor = RGB(0, 0, 0) 'колір переднього плану
    obj.BackStyle = fmBackStyleTransparent 'тип фону
Next obj
'метод SetFocus
CommandButton1.SetFocus 'установити фокус на кнопці
Debug.Print UserForm2.ActiveControl.Value 'значення активного елемента, який містить фокус
End Sub
