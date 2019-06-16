Dim x As Double
Dim s As String

'Процедура обробки події Initialize (ініціалізація форми)
Private Sub UserForm_Initialize()
TextBox1.AutoSize = False 'авторозмір
TextBox1.TextAlign = fmTextAlignCenter 'вирівнювання тексту
TextBox1.Enabled = False 'доступність
TextBox1.Text = "Hello" 'текст
TextBox1.Text = 5.71 'текст
x = CDbl(TextBox1.Text) 'конвертувати текст в дійсне число
'або
x = CDbl(TextBox1.Value) 'значення
TextBox1.SelStart = 0 'початкова позиція виділення
TextBox1.SelLength = TextBox1.TextLength 'довжина виділення
TextBox1.Copy 'скопіювати в буфер обміну
TextBox2.MaxLength = 8 'максимальна довжина тексту
TextBox2.PasswordChar = "*" 'символ для введення пароля
TextBox3.MultiLine = True 'багаторядковий режим
TextBox3.Height = 50 'висота
TextBox3.Font.Size = 12 'розмір шрифту
TextBox3.ScrollBars = fmScrollBarsBoth 'смуги прокручування
TextBox3.TabIndex = 2 'порядок зміни фокусу клавішею Tab
TextBox3.TabKeyBehavior = True 'дозволити вводити у текст табуляцію клавішею Tab
'присвоїти текст у двох рядках
TextBox3.Text = "Перший рядок" & Chr(13) & "Другий рядок"
TextBox3.SetFocus 'установити фокус
End Sub

'процедура обробки події MouseUp (відпущено кнопку миші)
Private Sub TextBox3_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
s = TextBox3.SelText 'присвоїти виділений текст
End Sub
'процедура обробки події MouseDown (натиснуто кнопку миші)
Private Sub TextBox4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
TextBox4.Text = s 'присвоїти текст
End Sub

'процедура обробки події Change (зміна значення)
Private Sub TextBox4_Change()
TextBox1.Text = TextBox4.Text 'присвоїти текст
End Sub

'процедура обробки події KeyPress (натиснута клавіша на клавіатурі)
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'якщо ASCII код клавіші <48 або >57 (не цифра)
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0 'не виводити нічого
End If
End Sub
