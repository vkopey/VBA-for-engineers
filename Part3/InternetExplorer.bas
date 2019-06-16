Dim IE As Object 'об'єкт програма Internet Explorer

Public Sub main()
Set IE = CreateObject("InternetExplorer.Application") 'створити об'єкт
IE.Navigate "About:Blank" 'початкова сторінка
IE.Toolbar = False 'відключити панель інструментів
IE.StatusBar = False 'відключити рядок стану
Do
Loop While IE.Busy 'цикл "поки браузер зайнятий"
IE.Visible = True 'зробити видимим
'заголовок html документу
IE.Document.Write "<html><title>My table</title><body>"
IE.Document.Write "<p>Table 1</p>" 'абзац
IE.Document.Write "<table border=1>" 'таблиця
'заголовок таблиці
IE.Document.Write "<tr><td><b>N</b></td><td><b>Name</b></td></tr>"
IE.Document.Write "<tr><td>1</td><td>Vasya</td></tr>" 'перший рядок
IE.Document.Write "<tr><td>2</td><td>Vova</td></tr>" 'другий рядок
IE.Document.Write "</table></body></html>" 'кінець таблиці і документа
End Sub
