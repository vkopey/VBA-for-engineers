Dim x As Double
Dim s As String
Dim b As Boolean
Dim d As Date
Dim i As Integer

Public Sub main()
x = Val("3.51") 'конвертує рядок в число, x=3.51
s = Str(x) 'конвертує число в рядок, s="3.51"
s = Format(Date, "ddd, d mm yyyy") 'форматує дані, s="Вт, 30 09 2008"
s = Format(Date, "dddd, d mm yyyy") 's="вторник, 30 09 2008"
s = Format(15790.335, "##,##0.00 грн") 's="15 790,34 грн" (тип Variant)
s = Format$(15790.335, "##,##0.00 грн") 's="15 790,34 грн" (тип String)
b = CBool("false") 'конвертує в булеве, b=false
d = CDate("30.09.2008") 'конвертує в дату, d=#30.09.2008#
x = CDbl("3,51") 'конвертує в дійсне, x=3.51
i = CInt("15") 'конвертує  в ціле, i=15
s = CStr(3.51) 'конвертує в рядок, s="3.51"
y = CVar(15 & "00") 'конвертує у Variant, y=1500
End Sub
