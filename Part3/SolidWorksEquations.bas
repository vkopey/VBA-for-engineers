Option Explicit
Sub main()
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swEqnMgr As SldWorks.EquationMgr 'менеджер рівнянь
    Dim i As Long 'індекс рівняння
    Dim nCount As Long 'кількість рівнянь

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc 'активний документ
    Set swEqnMgr = swModel.GetEquationMgr 'менеджер рівнянь
    Debug.Print "File = " & swModel.GetPathName 'шлях до моделі
    nCount = swEqnMgr.GetCount 'кількість рівнянь
    swEqnMgr.Equation(0) = """h"" = 10.0" 'змінити перше
    For i = 0 To nCount - 1 'для кожного рівняння вивести інформацію
        Debug.Print "  Equation(" & i & ")  = " & swEqnMgr.Equation(i)
        Debug.Print "    Value = " & swEqnMgr.Value(i)
        Debug.Print "    Index = " & swEqnMgr.Status
        Debug.Print "    Global variable? " & swEqnMgr.GlobalVariable(i)
    Next i
End Sub
