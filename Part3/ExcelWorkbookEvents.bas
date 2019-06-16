'Процедура обробки події Open (відкриття книги)
Private Sub Workbook_Open()
MsgBox "Збірник прикладів VBA-програм" & Chr(13) & "Автор: Копей В.Б."
End Sub

'Процедура обробки події BeforeClose (перед закриттям книги)
Private Sub Workbook_BeforeClose(Cancel As Boolean)
MsgBox "VBA - найлегше програмування!"
End Sub

'Процедура обробки події SheetSelectionChange (зміна виділення на листі)
Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)
Debug.Print "Зміна виділення " & Target.Address
End Sub

'Процедура обробки події SheetChange (зміна на листі)
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal Target As Range)
Debug.Print "Зміна у " & Target.Address
End Sub

'Процедура обробки події SheetBeforeRightClick (натиск правої кнопки міші на листі)
Private Sub Workbook_SheetBeforeRightClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
Debug.Print "Правий натиск на " & Target.Address
End Sub
