Dim obj As Class1 'оголосити об'єкт класу Class1

Public Sub main()
Set obj = New Class1 'створити об'єкт obj
obj.z = 1 'присвоїти властивості 'z' значення
obj.x = 2 'присвоїти властивості 'x' значення
Debug.Print obj.y(5) 'викликати метод 'y' з параметром 5
End Sub
