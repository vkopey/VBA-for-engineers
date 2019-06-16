Dim arg As Variant 'поле класу (закрите)
Public z As Integer 'поле класу (загальнодоступне)
Public Event notNumber(x) ' подія (загальнодоступна)

Private Sub Class_Initialize() 'процедура ініціалізації
arg = 0
End Sub

'процедура повернення властивістю 'x' значення дійсного типу
Public Property Get x() As Variant
x = arg 'присвоїти властивості значення
End Property

'процедура присвоєння властивості 'x' значення дійсного типу
Public Property Let x(ByVal vNewValue As Variant)
If Not IsNumeric(vNewValue) Then 'якщо не числове дане, то
    RaiseEvent notNumber(vNewValue) ' викликати подію
    Exit Property ' вийти
End If
arg = vNewValue 'присвоїти arg нове значення
End Property

'метод класу (функція) 'y' з параметром 's'
Public Function y(s As Variant)
y = z * arg ^ s
End Function
