Dim d, T As Date
Dim i As Integer
Dim l As Long
Dim f As Single

Public Sub main()
d = Date 'поточна дата, d=#9/30/2008#
T = Time 'поточний час, d=#20:31:07#
d = Now 'поточні дата і час, d=#9/30/2008 20:31:07#
d = #9/30/2008# 'присвоїти дату
T = #8:31:07 PM# 'присвоїти час
d = #9/30/2008 8:31:07 PM# 'присвоїти дату і час
d = DateSerial(2008, 9, 30) 'дата, задана цілими числами
T = TimeSerial(20, 31, 7) 'час, заданий цілими числами
i = Day(d) 'день
i = Month(d) 'місяць
i = Year(d) 'рік
i = Hour(T) 'година
i = Minute(T) 'хвилина
i = Second(T) 'секунда
i = Weekday(T) 'день тижня
f = Timer 'число секунд після опівночі
l = DateDiff("h", #8/12/1978#, Now) 'кількість інтервалів між датами (годин)
i = DatePart("m", Date) 'компонент дати (місяць)
d = DateAdd("m", 10, Date) 'додає до дати інтервал (10 місяців)
T = TimeValue("3:31:30 PM") 'перетворює рядок в формат часу
End Sub
