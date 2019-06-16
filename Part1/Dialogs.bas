Dim x As Double
Dim i As Byte

Public Sub main()
x = InputBox("введіть x", "x", 0) 'виводить діалогове вікно з полем введення і повертає введене значення
i = MsgBox("x=" & x, vbYesNoCancel, "Аргумент") 'виводить діалогове вікно з повідомленням і кнопками Yes/No/Cancel
MsgBox (i) 'код натиснутої кнопки
End Sub
