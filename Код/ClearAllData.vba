Sub ClearAllData()
'
' Очистка всех листов с данными перед запуском обновления
'

Sheets("data").Select
Cells.Select
Selection.Clear
Cells(1, 1).Select

For i = 1 To 5

    Sheets(CStr(i)).Select
    Cells.Select
    Selection.Clear
    Cells(1, 1).Select

Next

End Sub