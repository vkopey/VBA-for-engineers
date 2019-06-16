Dim sh As IWshRuntimeLibrary.WshShell 'об'єкт WshShell

Public Sub main()
Set sh = CreateObject("WScript.Shell") 'створити об'єкт WshShell
'__________________________________________________
'Створення файлу сценарію мовою Jscript
Open "c:\my.js" For Output As #1 'відкрити файл для запису
Print #1, "var objArgs = WScript.Arguments;" 'записати у файл
Print #1, "for (i=0; i<=objArgs.Count()-1;i++)" 'записати у файл
Print #1, "{WScript.Echo(parseFloat(objArgs(i))+2);}" 'записати у файл
Close #1 'закрити файл
sh.Run "wscript.exe c:\my.js 1 2" 'виконання сценарію з параметрами
'__________________________________________________
'Створення файлу сценарію мовою VBscript
Open "c:\my.vbs" For Output As #1 'відкрити файл для запису
Print #1, "For Each obj In WScript.Arguments" 'записати у файл
Print #1, "x=CDbl(WScript.StdIn.Readline)" 'записати у файл
Print #1, "WScript.StdOut.WriteLine obj+x" 'записати у файл
Print #1, "Next" 'записати у файл
Close #1 'закрити файл
sh.Run "cscript.exe c:\my.vbs 1 2" 'виконання в консолі сценарію з параметрами
'__________________________________________________
'Створення файлу сценарію XML мовами VBscript і Jscript
Open "c:\my.wsf" For Output As #1 'відкрити файл для запису
Print #1, "<package>"
Print #1, "<job id=""myjob"">" 'один пакет може містити кілька робіт
Print #1, "<script language=""VBScript"">"
Print #1, "Function F"
Print #1, "WScript.Echo WScript.ScriptFullName"
Print #1, "End Function"
Print #1, "</script>"
Print #1, "<script language=""JScript"">"
Print #1, "F();" 'виклик функції, написаній на VBscript
Print #1, "</script>"
Print #1, "</job>"
Print #1, "</package>"
Close #1 'закрити файл
sh.Run "wscript.exe //job:""myjob"" c:\my.wsf" 'виконання сценарію
End Sub
