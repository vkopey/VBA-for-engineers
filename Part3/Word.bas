Public Sub main()
Dim wdApp As Word.Application 'об'єкт програма Word 
Dim wdDoc As Word.Document 'об'єкт документ
Set wdApp = CreateObject("Word.Application") 'створити об'єкт Word
Set wdDoc = wdApp.Documents.Add 'створити об'єкт документ
wdApp.Visible = True 'зробити видимим Word
wdDoc.ActiveWindow.View.Zoom.PageFit = wdPageFitBestFit 'масштаб виду
wdDoc.ActiveWindow.Selection.TypeText "hello world!" 'надрукувати текст
wdDoc.ActiveWindow.Selection.TypeParagraph 'надрукувати знак абзацу
wdDoc.ActiveWindow.Selection.TypeText "Visual Basic for Applications"
'колір діапазону з перших шести символів
wdDoc.Range(0, 6).Font.Color = wdColorBlue
'виділити діапазон з перших двох слів
wdDoc.Range(wdDoc.Words(1).Start, wdDoc.Words(2).End).Select
'розширити виділення до символу "!"
wdDoc.ActiveWindow.Selection.Extend Character:="!"
'вивести текст виділення
Debug.Print wdApp.ActiveDocument.ActiveWindow.Selection.Text
wdDoc.ActiveWindow.Selection.Copy 'копіювати виділення в буфер обміну
wdDoc.ActiveWindow.Selection.InsertAfter "!!" 'вставити після виділення "!!"
'відмінити виділення і перемістити курсор в його кінець
wdDoc.ActiveWindow.Selection.Collapse Direction:=wdCollapseEnd
wdApp.Documents.Add 'додати документ
n = wdApp.ActiveDocument.name 'ім'я активного документа
wdApp.ActiveWindow.Selection.Paste 'вставити з буфера обміну
wdApp.Documents(2).Windows(1).Activate 'активувати другий документ
'надати жирний шрифт першому слову першого абзацу
wdApp.Documents(2).Paragraphs(1).Range.Words(1).Font.Bold = True
'вивести перший символ першого слова першого речення
Debug.Print wdApp.Documents(2).Sentences(1).Words(1).Characters(1)
wdApp.Documents(n).SaveAs ("e:\my_doc.doc") 'зберегти як
wdApp.Documents("my_doc").Close 'закрити документ з іменем my_doc
wdApp.Documents.Open ("e:\my_doc.doc") 'відкрити документ
'знайти слова "world" і замінити на "World"
With ActiveDocument.Content.Find
    .ClearFormatting
    .Text = "world"
    .Replacement.ClearFormatting
    .Replacement.Text = "World"
    .Execute Replace:=wdReplaceAll
End With
wdApp.Quit 'вийти з Word
End Sub
