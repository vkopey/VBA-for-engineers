Dim d As New Scripting.Dictionary 'словник
Dim k As Variant

Public Sub main()
'додати в словник елемент з ключем 1 і значенням "Січень"
d.Add 1, "Січень"
d.Add 2, "Лютий"
d.Add 3, "Березень"
For Each k In d 'для кожного ключа 'k' в словнику 'd'
    Debug.Print k, d(k) 'вивести ключ і значення
    'Debug.Print k, d.Item(k) 'або так
Next k
d.Remove 2 'видалити елемент з ключем 2
Debug.Print d.Exists(2) 'чи існує елемент із ключем 2
d.Key(3) = 2 'змінити ключ
Debug.Print d.Count 'вивести кількість елементів
Debug.Print d.keys(0), d.items(0) 'перші елементи масивів ключів і значень
Set d = Nothing
End Sub
