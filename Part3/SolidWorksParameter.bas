Public Sub main()
Dim swApp As Object 'об'єкт SolidWorks Application
Dim Part As Object 'об'єкт документ SolidWorks
Set swApp = CreateObject("SldWorks.Application") 'створити об'єкт
Set Part = swApp.ActiveDoc 'активний документ
'змінити значення параметра на 20 мм
Part.Parameter("D1@Extrude1").SystemValue = 20 / 1000
Part.EditRebuild 'перебудувати модель
End Sub
