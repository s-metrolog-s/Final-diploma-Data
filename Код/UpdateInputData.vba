Sub UpdateInputData()
'
' Загрузка и обновление данных по продажам и марже
'
    ' Загрузка данных из файла data
    ' файл располагается в той же папке, что и сам макрос
    ' --------------------------------------------------------------------------
    
    'открываем файл
    Workbooks.Open Filename:=ThisWorkbook.Path & "\data.xlsx"
    Cells.Select
    Selection.Copy
    
    'возвращаемся в исходный
    Windows("ProjectX.xlsm").Activate
    Sheets("data").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'очистка буфера обмена
    Application.CutCopyMode = False
    
    'особенность входных данных - первые две ячейки пустые
    'проверяем, что все так и есть, и удаляем
    If Cells(1, 2).Value = "" And Cells(2, 2).Value = "" Then
        
        Rows("1:2").Select
        Selection.Delete Shift:=xlUp
        
    End If
    
    'особенность входных данных - вторая ячейка с данными суммирующая
    'проверяем, что все так и есть, и удаляем
    If Cells(2, 1).Value = "Итого" Then
    
        Rows("2:2").Select
        Selection.Delete Shift:=xlUp
        
    End If
    
    Windows("data.xlsx").Activate
    Windows("data.xlsx").Close
    ' --------------------------------------------------------------------------

    Windows("ProjectX.xlsm").Activate
    Sheets("Settings").Select

End Sub