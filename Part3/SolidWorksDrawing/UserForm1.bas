Dim swApp As SldWorks.SldWorks
Dim view As New DrawingView
Dim data As New DrawingData

Private Sub ComboBox1_Change() 'вибрано лист
ComboBox2.Clear
For Each s In data.Sheets
    If s.Name = ComboBox1.Text Then
        For Each v In s.views
            ComboBox2.AddItem v.Name
        Next v
    End If
Next s
End Sub

Private Sub ComboBox2_Change() 'вибрано вигляд
Set view = findView(ComboBox1.Text, ComboBox2.Text, data)
writeViewFields
End Sub

Private Sub CommandButton2_Click() 'кнопка save
Set view = findView(ComboBox1.Text, ComboBox2.Text, data)
readViewFields
End Sub

Private Sub CommandButton1_Click() 'кнопка GenerateDrawing
GenerateDrawing data
End Sub

Private Sub CommandButton3_Click() 'кнопка UpdateDrawing
UpdateDrawing TextBox6.Text, data
End Sub

Private Sub UserForm_Initialize() 'ініціалізація форми
Dim view1 As New DrawingView
Dim view2 As New DrawingView
Dim sheet1 As New DrawingSheet
Dim sheet2 As New DrawingSheet

Set swApp = Application.SldWorks
macroPath = UCase(swApp.GetCurrentMacroPathFolder)
Macro21.macroPath = macroPath 'шлях до каталогу з макросом
Debug.Print macroPath

'заповнити дані
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

sheet1.views.Add view1
sheet1.views.Add view2
data.Sheets.Add sheet1
data.Sheets.Add sheet2

'заповнити список листів
For Each s In data.Sheets
    ComboBox1.AddItem s.Name
Next s

TextBox6.Text = macroPath & "\Drawing1.SLDDRW"

End Sub

Public Sub writeViewFields() 'записати дані поточного виду в поля
ComboBox2.Text = view.Name
TextBox1.Text = view.ModelViewName
TextBox2.Text = view.PositionX
TextBox3.Text = view.PositionY
TextBox4.Text = view.ReferenceModelPath
TextBox5.Text = view.Configuration
End Sub

Public Sub readViewFields() 'прочитати дані поточного виду з полів
view.Name = ComboBox2.Text
view.ModelViewName = CInt(TextBox1.Text)
view.PositionX = CDbl(TextBox2.Text)
view.PositionY = CDbl(TextBox3.Text)
view.ReferenceModelPath = TextBox4.Text
view.Configuration = TextBox5.Text
End Sub
