Dim con As ADODB.Connection 'з'єднання
Dim rs As ADODB.Recordset 'набір записів

Public Sub main()
Set con = New ADODB.Connection 'створити з'єднання
con.Provider = "Microsoft.Jet.OLEDB.4.0" 'провайдер
con.ConnectionString = "db1.mdb" 'рядок з'єднання до джерела даних
con.Open 'відкрити з'єднання
Set rs = New ADODB.Recordset 'створити набір записів
rs.Source = "SELECT * FROM stud WHERE (Оцінка>3)" 'рядок запиту SQL
Set rs.ActiveConnection = con 'активне з'єднання
rs.Open 'відкрити набір записів
rs.MoveFirst 'перейти до першого запису
Do While Not rs.EOF 'поки не кінець записів
    Debug.Print rs.Fields(0).Value 'вивести значення першого поля
    Debug.Print rs.Fields(1).Value 'вивести значення другого поля
    rs.MoveNext 'перейти до наступного запису
Loop
rs.Save "db2.xml", adPersistXML 'зберегти як файл XML
rs.Close 'закрити набір записів
con.Close 'закрити з'єднання
End Sub
