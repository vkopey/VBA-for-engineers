Dim x, y As Double
Dim A, b, c As Boolean

Public Sub main()
x = 2: y = 3
A = True: b = False
c = x > y 'більше: c=False
c = x < y 'менше: c=True
c = x >= y 'більше дорівнює: c=False
c = x <= y 'менше дорівнює: c=True
c = x <> y 'не дорівнює: c=True
c = x = y 'дорівнює: c=False
c = A And b 'логічне "І": c=False
c = A Or b 'логічне "АБО": c=True
c = Not A 'логічне "НЕ": c=False
c = A Xor b 'виключна диз’юнкція: c=True
End Sub
