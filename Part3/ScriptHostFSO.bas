Dim fso As IWshRuntimeLibrary.FileSystemObject 'об'єкт файлової системи
Dim drvs As IWshRuntimeLibrary.Drives 'диски
Dim drv As IWshRuntimeLibrary.Drive 'диск
Dim fl As IWshRuntimeLibrary.File 'файл
Dim ts As IWshRuntimeLibrary.TextStream 'текстовий потік

Public Sub main()
'об'єкт файлової системи
Set fso = CreateObject("Scripting.FileSystemObject")
Set drvs = fso.Drives 'диски
Debug.Print drvs.Count 'вивести кількість дисків
For Each drv In drvs 'для кожного диску
 'якщо тип диску Removable і він доступний, тоді
 If drv.DriveType = Removable And drv.IsReady Then
    Debug.Print drv.DriveLetter 'вивести букву диска
    'для кожного файлу в кореневій директорії
    For Each fl In drv.RootFolder.Files
        If fl.name = "Autorun.inf" Then 'якщо ім'я "Autorun.inf", то
            fl.Copy "d:\" 'копіювати на диск d
            'відкрити файл для читання
            Set ts = fl.OpenAsTextStream(ForReading)
            Debug.Print ts.ReadAll 'читати все
            ts.Close 'закрити
            fl.name = "Autorun_.inf" 'змінити ім'я
        End If
    Next
 End If
Next
End Sub
