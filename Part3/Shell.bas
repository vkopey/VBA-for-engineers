Dim shl As Shell32.Shell 'об'єкт Shell

Public Sub main()
Set shl = CreateObject("Shell.Application") 'створити об'єкт Shell
shl.ControlPanelItem "appwiz.cpl" 'відкрити "Установка і видалення програм"
shl.Explore "c:\" 'відкрити Explorer
shl.Open "c:\windows" 'відкрити папку
shl.FileRun 'відкрити "Запуск програми"
Set Folder = shl.Namespace("d:\") 'створити об'єкт папка d:\
Folder.CopyHere "c:\boot.ini" 'копіювати в d:\ файл
Set File = Folder.parsename("boot.ini") 'створити об'єкт файл boot.ini
File.InvokeVerb "Open" 'відкрити файл
End Sub
