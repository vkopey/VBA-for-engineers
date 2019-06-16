Public Sub main()
Worksheets(1).Activate 'активувати перший лист
ActiveSheet.Visible = True 'зробити видимим активний лист
ActiveSheet.UsedRange.Clear 'очистити діапазон з значеннями
Debug.Print ActiveSheet.StandardHeight 'стандартна висота рядків
Debug.Print Worksheets.Count 'кількість листів
Worksheets.Add 'додати лист на початок
s = Worksheets(1).name 'присвоїти ім'я першого листа
Worksheets(1).Copy Worksheets(3) 'скопіювати перед третім листом
Worksheets(1).Move Worksheets(3) 'перемістити перед третім листом
Worksheets(s).Delete 'знищити
Worksheets(2).Delete 'знищити
Worksheets(2).Activate 'активувати другий лист
Worksheets("Лист1").Activate 'активувати лист "Лист1"
End Sub
