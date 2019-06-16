Dim swApp As SldWorks.SldWorks 'об'єкт програма-SolidWorks
Dim swModel As SldWorks.ModelDoc2 'об'єкт модель
Dim swDraw As SldWorks.DrawingDoc 'об'єкт креслення
Dim myView As SldWorks.view 'об'єкт вид
Dim swModelDocExt As SldWorks.ModelDocExtension 'об'єкт розширення моделі
Dim swDrawingComponent As SldWorks.DrawingComponent 'об'єкт компонент креслення
Dim swSelectionMgr As SldWorks.SelectionMgr 'об'єкт менеджера вибору

Public macroPath As String 'шлях до каталогу з макросом (робочого каталогу)
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long

'процедура створює креслення за даними data
Sub GenerateDrawing(data As DrawingData)
'масив імен стандартних видів
ViewNames = Array("*Front", "*Front", "*Back", "*Left", "*Right", "*Top", "*Bottom")

Set swApp = Application.SldWorks
'нове креслення
Set swDraw = swApp.NewDocument(macroPath & "\\a3 - gost_sh1.slddrt", 0, 0, 0)
For Each s In data.Sheets 'для кожного листа в data
    'новий лист
    boolstatus = swDraw.NewSheet3(s.Name, 12, 12, 2, 1, True, macroPath & "\\a3 - gost_sh1.slddrt", 0.42, 0.297, "По умолчанию")
    For Each v In s.views 'для кожного виду листа в data
        'новий вид
        Set myView = swDraw.CreateDrawViewFromModelView3(v.ReferenceModelPath, ViewNames(v.ModelViewName), v.PositionX, v.PositionY, 0)
        'задати конфігурацію
        myView.ReferencedConfiguration = v.Configuration 
    Next v
Next s

End Sub

'оновлює креслення drawingPathName за даними data
Public Sub UpdateDrawing(drawingPathName As String, data As DrawingData)
Dim sname As String
Dim compName As String

Set swApp = Application.SldWorks
'відкрити креслення
Set swDraw = swApp.OpenDoc6(drawingPathName, 3, 0, "", longstatus, longwarnings)

snames = swDraw.GetSheetNames 'назви листів креслення
For i = 0 To swDraw.GetSheetCount - 1 'для кожного листа
    sname = snames(i) 'назва листа
    swDraw.ActivateSheet sname 'активувати лист
    Debug.Print sname
    
    vs = swDraw.Sheet(sname).GetViews 'види листа
    If TypeName(vs) <> "Empty" Then 'якщо є види
        For Each v In vs 'для кожного виду
            swDraw.ActivateView v.Name 'активувати вид
            Debug.Print v.Name, v.GetReferencedModelName 'назва виду і його вихідної моделі
            compName = ""
            If TypeName(v.GetVisibleComponents) <> "Empty" Then
                For Each c In v.GetVisibleComponents 'для кожного видимого компонента
                    compName = c.Name 'назва компонента
                Next c
            End If
            pos = v.Position 'позиція виду
            Set myView = v 'об'єкт виду (необхідно для безпомилкової роботи функції GetOrientationName)
            Debug.Print myView.GetOrientationName 'назва орієнтації виду (*Front, *Left...)
            
            Set dv = findView(sname, v.Name, data) 'знайти вид у data
            If Not dv Is Nothing Then 'якщо є
                Debug.Print v.GetReferencedModelName, dv.ReferenceModelPath
                'якщо назви вихідної моделі не однакові
                If v.GetReferencedModelName <> dv.ReferenceModelPath Then
                    Debug.Print "ReplacingViewModel..."
                    'замінити вихідну модель на нову
                    ReplaceViewModel sname, v.Name, compName, dv.ReferenceModelPath
                End If
                pos(0) = dv.PositionX
                pos(1) = dv.PositionY
                v.Position = pos 'нова позиція
                v.ReferencedConfiguration = dv.Configuration 'нова конфігурація
                
                'вибрати вид
                boolstatus = swDraw.Extension.SelectByID2(v.Name, "DRAWINGVIEW", 0, 0, 0, False, 0, Nothing, 0)
                swDraw.ShowNamedView2 "", dv.ModelViewName 'нова орієнтація виду, або swDraw.ShowNamedView2 "*Right", -1
            End If
        Next v
    End If
Next i

swDraw.EditRebuild 'перебудувати модель
End Sub

'функція шукає вид у data
Public Function findView(sheetName As String, viewName As String, data As DrawingData) As DrawingView
Set findView = Nothing
For Each s In data.Sheets 'для кожного листа
    If s.Name = sheetName Then
        For Each v In s.views 'для кожного виду
            If v.Name = viewName Then
                Set findView = v 'функція повертає вид або Nothing
            End If
        Next v
    End If
Next s
End Function

'процедура заміни вихідної моделі у виді креслення на нову modelFileName
Public Sub ReplaceViewModel(sheetName As String, viewName As String, componentName As String, modelFileName As String)
Dim views(0) As Object 'вид
Dim instances(0) As Object 'компонент
Set swModel = swApp.ActiveDoc 'креслення
Set swModelDocExt = swModel.Extension
vs = swDraw.Sheet(sheetName).GetViews 'усі види
For Each v In vs 'для кожного виду
    If v.Name = viewName Then 'знайти вид з іменем viewName
        Set views(0) = v
    End If
Next v
id_ = componentName & "@" & viewName
'вибрати компонент
boolstatus = swModelDocExt.SelectByID2(id_, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
Set swSelectionMgr = swModel.SelectionManager
Set swDrawingComponent = swSelectionMgr.GetSelectedObject6(1, -1)
Set instances(0) = swDrawingComponent.Component
'замінити модель
boolstatus = swDraw.ReplaceViewModel(modelFileName, (views), (instances))
End Sub

Public Sub main() 'процедура для тестування GenerateDrawing та UpdateDrawing
'створити об'єкти
Dim view1 As New DrawingView
Dim view2 As New DrawingView
Dim sheet1 As New DrawingSheet
Dim sheet2 As New DrawingSheet
Dim data As New DrawingData

Set swApp = Application.SldWorks
macroPath = UCase(swApp.GetCurrentMacroPathFolder) 'шлях до каталогу з макросом
Debug.Print macroPath

'присвоїти атрибутам значення
view1.Name = "Drawing View1"
view2.Name = "Drawing View2"
view1.ReferenceModelPath = macroPath & "\part1.SLDPRT"
view2.ReferenceModelPath = macroPath & "\part1.SLDPRT"
view1.Configuration = "По умолчанию"
view2.Configuration = "По умолчанию"
view1.ModelViewName = 0
view2.ModelViewName = 3
sheet1.Name = "Sheet0"
sheet2.Name = "Sheet1"
view1.PositionX = 0.15
view1.PositionY = 0.19
view2.PositionX = 0.3
view2.PositionY = 0.19

sheet1.views.Add view1 'додати види в листи
sheet1.views.Add view2
data.Sheets.Add sheet1 'додати листи в data
data.Sheets.Add sheet2

'Set dv = findView("Sheet0", "Drawing View1", data) ' тест функції пошуку
'If Not dv Is Nothing Then Debug.Print dv.ReferenceModelPath

'вікно з кнопками Yes-No
nvar = MsgBox("YES for Generate Drawing" & Chr(13) & "or NO for Update Drawing", 4, "Press")
Debug.Print nvar
If nvar = 6 Then 'якщо вибрано Yes
    GenerateDrawing data 'створити креслення
Else
    'інакше змінити атрибути і оновити креслення
    view1.ReferenceModelPath = macroPath & "\part2.SLDPRT"
    view2.ReferenceModelPath = macroPath & "\part2.SLDPRT"
    view2.Configuration = "20"
    view1.ModelViewName = 2
    view1.PositionX = 0.14
    view1.PositionY = 0.18
    UpdateDrawing macroPath & "\Drawing1.SLDDRW", data 'оновити
End If
End Sub
