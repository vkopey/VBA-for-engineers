Public Sub main()
'Додати дані на лист:
Range("A1:A6")=Application.Transpose(Array(1,2,3,4,5,6))
Range("B1:B6")=Application.Transpose(Array(0,3,6,7,9,11))
Charts.Add 'додати діаграму
ActiveChart.ChartType = xlXYScatterSmooth 'тип діаграми
ActiveChart.SetSourceData Source:=Sheets("Лист1").Range("A1:B6"), PlotBy:= _
    xlColumns 'дані для діаграми
ActiveChart.Location Where:=xlLocationAsObject, Name:="Лист1" 'розміщення
With ActiveChart 'з активною діаграмою
    .HasTitle = True 'має заголовок
    .ChartTitle.Characters.Text = "Графік" 'надпис заголовка
    .Axes(xlCategory, xlPrimary).HasTitle = True 'має надпис осі категорій
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "x" 'надпис
    .Axes(xlValue, xlPrimary).HasTitle = True 'має надпис осі значень
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "y" 'надпис
    'додати криву регресії, її рівняння, та величину достовірності апроксимації
    .SeriesCollection(1).Trendlines.Add Type:=xlLogarithmic, _
    DisplayEquation:=True, DisplayRSquared:=True
End With
End Sub
