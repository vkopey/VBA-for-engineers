Dim РобочаОбласть As DAO.Workspace 'робоча область
Dim БазаДаних As DAO.Database 'база даних
Dim Таблиця As DAO.TableDef 'таблиця
Dim Поле As DAO.Field 'поле
Dim Запис As DAO.Recordset 'набір записів
Dim Sqlstr As String 'рядок для команд SQL

Public Sub main()
'Створюємо робочу область
Set РобочаОбласть = CreateWorkspace("", "admin", "", dbUseJet)
'Створюємо базу даних "db1.mdb"
Set БазаДаних = РобочаОбласть.CreateDatabase("db1", dbLangGeneral)
Set Таблиця = New DAO.TableDef 'створюємо таблицю
With Таблиця 'з таблицею
    .Fields.Append .CreateField("Прізвище", dbText) 'створити поле "Прізвище"
    .Fields.Append .CreateField("Оцінка", dbInteger) 'створити поле "Оцінка"
End With
Таблиця.name = "stud" 'ім'я таблиці
БазаДаних.TableDefs.Append Таблиця 'додати таблицю
БазаДаних.Close 'закрити базу даних
Set БазаДаних = РобочаОбласть.OpenDatabase("db1.mdb", True) 'відкрити базу даних
'Створюємо записи таблиці "stud"
Set Запис = БазаДаних.OpenRecordset("stud", dbOpenDynaset)
Запис.AddNew 'додати запис
Запис.Fields("Прізвище").Value = "Іванов" 'ввести значення поля "Прізвище"
Запис.Fields("Оцінка").Value = 4 'ввести значення поля "Оцінка"
Запис.Update 'оновити
Запис.AddNew 'додати запис
Запис.Fields("Прізвище").Value = "Петров"
Запис.Fields("Оцінка").Value = 3
Запис.Update 'оновити
Запис.Close 'Закрити записи

'Команда SQL вибирає записи з таблиці stud, де Оцінка>3
Sqlstr = "SELECT * FROM stud WHERE (Оцінка>3)"
'Створюємо записи за запитом SQL
Set Запис = БазаДаних.OpenRecordset(Sqlstr, dbOpenDynaset)
'Виведення записів запиту
Запис.MoveFirst 'перейти на перший запис
Do While Not Запис.EOF 'поки не кінець записів
Debug.Print Запис.Fields("Прізвище").Value 'вивести значення поля "Прізвище"
Debug.Print Запис.Fields("Оцінка").Value 'вивести значення поля "Оцінка"
Запис.MoveNext 'перейти на наступний запис
Loop
'Закрити запис, базу даних і робочу область
Запис.Close
БазаДаних.Close
РобочаОбласть.Close
End Sub
