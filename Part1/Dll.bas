'функція MessageBeep з системної бібліотеки user32.dll
Declare Sub MessageBeep Lib "user32.dll" (ByVal T As Long)
'функція PlaySound з системної бібліотеки winmm.dll
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
(ByVal n As String, ByVal m As Long, ByVal f As Long) As Long

Public Sub main()
MessageBeep 32 'виклик функції з параметром
'виклик функції з параметром
Call PlaySound("c:\WINDOWS\Media\tada.wav", 0&, SND_ASYNC Or SND_FILENAME)
End Sub
