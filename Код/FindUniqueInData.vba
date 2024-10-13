Sub FindUniqueInData()
'
' Создание списков уникальных значений для формирования результирующей таблицы
'

Dim Col_number_of_stores As Integer
Dim Col_ts_fio As Integer
Dim Col_articles As Integer
Dim Col_sales As Integer

' Использование глобальных переменных отменено по ТЗ, реализована возможность настройки
' пользователем номеров столбцов во входном файле, так как у части учетных записей
' данные могут выгружаться в другом порядке, загружаем с листа Settings

Col_number_of_stores = Sheets("Settings").Cells(2, 6).Value
Col_ts_fio = Sheets("Settings").Cells(3, 6).Value
Col_articles = Sheets("Settings").Cells(4, 6).Value
Col_sales = Sheets("Settings").Cells(5, 6).Value
Col_branch = Sheets("Settings").Cells(6, 6).Value
Col_sub_branch = Sheets("Settings").Cells(7, 6).Value

Sheets("data").Select
'находим последнюю непустую строку на листе
LastRow = Cells(Rows.Count, 2).End(xlUp).Row

'-------------------------------------------------------------------------------------------------
'Формируем список филиалов

'выделяем диапазон всех ячеек с дублирующейся информацией
Range(Cells(2, Col_number_of_stores), Cells(LastRow, Col_number_of_stores)).Select
Selection.Copy
Sheets("1").Select
Cells(1, 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
LastRow_temp = Cells(Rows.Count, 1).End(xlUp).Row

'оставляем только уникальные значения
ActiveSheet.Range(Cells(1, 1), Cells(LastRow_temp, 1)).RemoveDuplicates Columns:=1, Header:=xlNo
Cells(1, 1).Select

'--------------------------------
'Сортировка списка по возрастанию
    
    Cells.Select
    ActiveWorkbook.Worksheets("1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("1").Sort.SortFields.Add2 Key:=Range("A1:A65889"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("1").Sort
    'сортируем весь столбец
        .SetRange Range("A1:B65889")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
'--------------------------------

Sheets("data").Select
'-------------------------------------------------------------------------------------------------
'Формируем список фамилий менеджеров

'выделяем диапазон всех ячеек с дублирующейся информацией
Range(Cells(2, Col_number_of_stores), Cells(LastRow, Col_ts_fio)).Select
Selection.Copy
Sheets("2").Select
Cells(1, 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
LastRow_temp = Cells(Rows.Count, 1).End(xlUp).Row

'оставляем только уникальные значения
ActiveSheet.Range(Cells(1, 1), Cells(LastRow_temp, 2)).RemoveDuplicates Columns:=2, Header:=xlNo
Cells(1, 1).Select

'--------------------------------
'Сортировка списка по возрастанию
    
    Cells.Select
    ActiveWorkbook.Worksheets("2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("2").Sort.SortFields.Add2 Key:=Range("A1:A65889"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("2").Sort.SortFields.Add2 Key:=Range("B1:B65889"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("2").Sort
        .SetRange Range("A1:B65889")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
'--------------------------------

Sheets("data").Select
'-------------------------------------------------------------------------------------------------
'Формируем список артикулов

'выделяем диапазон всех ячеек с дублирующейся информацией
Range(Cells(2, Col_articles), Cells(LastRow, Col_articles + 1)).Select
Selection.Copy
Sheets("3").Select
Cells(1, 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
LastRow_temp = Cells(Rows.Count, 1).End(xlUp).Row

'оставляем только уникальные значения
ActiveSheet.Range(Cells(1, 1), Cells(LastRow_temp, 2)).RemoveDuplicates Columns:=1, Header:=xlNo
Cells(1, 1).Select

'--------------------------------
'Сортировка списка по возрастанию
    
    Cells.Select
    ActiveWorkbook.Worksheets("3").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("3").Sort.SortFields.Add2 Key:=Range("A1:A65889"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("3").Sort
        .SetRange Range("A1:B65889")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
    
'--------------------------------

Sheets("data").Select
'-------------------------------------------------------------------------------------------------
'Формируем список сегментов

Range(Cells(2, Col_branch), Cells(LastRow, Col_branch)).Select
Selection.Copy
Sheets("4").Select
Cells(1, 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
LastRow_temp = Cells(Rows.Count, 1).End(xlUp).Row
ActiveSheet.Range(Cells(1, 1), Cells(LastRow_temp, 2)).RemoveDuplicates Columns:=1, Header:=xlNo
Cells(1, 1).Select
Sheets("data").Select
'-------------------------------------------------------------------------------------------------
'Формируем список подсегментов

Range(Cells(2, Col_sub_branch), Cells(LastRow, Col_sub_branch)).Select
Selection.Copy
Sheets("5").Select
Cells(1, 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False
LastRow_temp = Cells(Rows.Count, 1).End(xlUp).Row
ActiveSheet.Range(Cells(1, 1), Cells(LastRow_temp, 2)).RemoveDuplicates Columns:=1, Header:=xlNo
Cells(1, 1).Select
Sheets("data").Select
'-------------------------------------------------------------------------------------------------

Sheets("Settings").Select


End Sub