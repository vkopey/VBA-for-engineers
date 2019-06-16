Dim V As SpeechLib.SpVoice 'голос

Public Sub main()
Set V = New SpeechLib.SpVoice 'новий голос
Set V.Voice = V.GetVoices("Name=Microsoft Sam", "Language=9").Item(0) 'параметри голосу
V.Speak "hello" 'сказати слово
End Sub
