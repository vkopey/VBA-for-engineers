Dim c As New Collection 'колекція
Dim x As Variant

Public Sub main()
c.Add "січень" 'додати елемент з ключем 1
c.Add "лютий" 'додати елемент з ключем 2
c.Add "грудень", before:=1  'вставити елемент перед 1
c.Add "березень", "5" 'додати елемент з ключем 5
c.Remove 2 'видалити елемент з ключем 2 ("січень")
Debug.Print c.Count 'кількість елементів
Debug.Print c(1), c(2), c("5") 'значення елементів за ключами
'Debug.Print c.Item("5")'або так
For Each x In c 'для всіх елементів 'x' в колекції 'c'
    Debug.Print x, 'вивести елемент
Next x
Set c = Nothing
End Sub
