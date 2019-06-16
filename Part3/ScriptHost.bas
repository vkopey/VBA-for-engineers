Dim sh As IWshRuntimeLibrary.WshShell 'об'єкт WshShell

Public Sub main()
Set sh = CreateObject("WScript.Shell") 'створення об'єкта WshShell
sh.Run "cmd.exe /c net share d=d:\" 'виконати cmd.exe з параметрами
sh.Run "cmd.exe /k net share d /delete"
app = Shell("calc.exe", 1) 'виконати calc.exe
sh.AppActivate "Calculator" 'активувати вікно (або так: AppActivate app)
sh.SendKeys "1{+}2{Enter}", True 'надіслати вікну клавіші
'прочитати з реєстру
Debug.Print sh.RegRead("HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaveTimeOut")
'записати в реєстр
sh.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\ScreenSaveTimeOut", _
"600", "REG_SZ"
Debug.Print sh.Popup("text", 0, "title", 1) 'показати діалогове вікно
Debug.Print sh.ExpandEnvironmentStrings("%WinDir%") 'шлях до папки Windows
sh.LogEvent 1, "Hello!" 'записує повідомлення з помилкою в системний журнал
End Sub
