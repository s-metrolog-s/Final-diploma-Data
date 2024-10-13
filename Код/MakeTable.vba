Sub Make_table()
'
' Формирование результирующей таблицы
'

'объявляем счетчик
Dim Counter As Integer

Application.CutCopyMode = False

'-------------------------------------------------------------------------------------------------
'Очистка листов Work и Tasks перед загрузкой новых данных

Sheets("Work").Select
Cells.Select
Selection.Clear
'убираем закрепление границ
ActiveWindow.FreezePanes = False
Cells.Select
Cells.FormatConditions.Delete
Cells(1, 1).Select

Sheets("Tasks").Select
Columns("C:C").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Delete Shift:=xlToLeft
Rows("4:4").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp
Cells.Select
'убираем условное форматирование со всего листа
Cells.FormatConditions.Delete
ActiveWindow.FreezePanes = False
Cells(1, 1).Select

'-------------------------------------------------------------------------------------------------
'Заполнение списка артикулов

Sheets("3").Select
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(1, 1), Cells(LastRow, 2)).Select
Selection.Copy
Sheets("Tasks").Select
Cells(4, 1).Select
ActiveSheet.Paste
Application.CutCopyMode = False

'-------------------------------------------------------------------------------------------------
'Заполнение списка всех менеджеров слева направо

Sheets("2").Select
Number_of_managers = Cells(Rows.Count, 1).End(xlUp).Row

'используем цикл для заполнения любого возможного количества менеджеров
Counter = 1
Do While Counter <= Number_of_managers

    Sheets("2").Select
    Cells(Counter, 2).Select
    Selection.Copy
    Sheets("Tasks").Select
    LastCol = Cells(3, Columns.Count).End(xlToLeft).Column + 1
    Cells(3, LastCol).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Counter = Counter + 1
    
Loop
'-------------------------------------------------------------------------------------------------
'Копируем номер филиала в привязке к менеджеру

LastCol_managers = Cells(3, Columns.Count).End(xlToLeft).Column

Cells(2, 2).Select
Selection.Copy
Range(Cells(2, 3), Cells(2, LastCol_managers)).Select
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False
        
Cells(2, 2).Select
Selection.Copy
Range(Cells(2, 3), Cells(3, LastCol_managers)).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Range(Cells(2, 3), Cells(2, LastCol_managers)).Select
Selection.Copy
Range(Cells(2, 3), Cells(2, LastCol_managers)).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False

'-------------------------------------------------------------------------------------------------
'Заполняем суммы продаж по каждому менеджеру, учитывая фильтрацию по сегментам и подсегментам
'
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'***************************************************************
'TO-DO отдать в тестирование, нормально ли по времени исполнения

Sheets("data").Select
LastRow_data = Cells(Rows.Count, 1).End(xlUp).Row
Sheets("Tasks").Select

Sum_column = Sheets("Settings").Cells(5, 6).Value
Article_column = Sheets("Settings").Cells(4, 6).Value
Manager_column = Sheets("Settings").Cells(3, 6).Value
Brunch_column = Sheets("Settings").Cells(6, 6).Value
SubBrunch_column = Sheets("Settings").Cells(7, 6).Value

Dim rng_1 As Range
Dim rng_2 As Range
Dim rng_3 As Range
Dim rng_4 As Range
Dim rng_5 As Range

Sheets("data").Select

'присваиваем необходимые диапазоны для удобства использования в циклах
Set rng_1 = Sheets("data").Range(Cells(1, Sum_column), Cells(LastRow_data, Sum_column))
Set rng_2 = Sheets("data").Range(Cells(1, Article_column), Cells(LastRow_data, Article_column))
Set rng_3 = Sheets("data").Range(Cells(1, Manager_column), Cells(LastRow_data, Manager_column))
Set rng_4 = Sheets("data").Range(Cells(1, Brunch_column), Cells(LastRow_data, Brunch_column))
Set rng_5 = Sheets("data").Range(Cells(1, SubBrunch_column), Cells(LastRow_data, SubBrunch_column))

Branch_Value = Sheets("Settings").Range("J2").Value
SubBranch_Value = Sheets("Settings").Range("J5").Value

Sheets("Tasks").Select

'---------------------
'Циклы заполнения в зависимости от условий выбора сегмента и подсегмента

Counter_sum = 1

If Branch_Value = "" And SubBranch_Value = "" Then

    For i = 3 To LastCol_managers
        
        For j = 4 To LastRow
            Cells(j, i).Value = WorksheetFunction.SumIfs(rng_1, rng_2, Cells(j, 1).Value, _
                rng_3, Cells(3, i).Value)
        Next
        Counter_sum = Counter_sum + 1
    Next
    
ElseIf Branch_Value = "" And SubBranch_Value <> "" Then

    For i = 3 To LastCol_managers
        
        For j = 4 To LastRow
            Cells(j, i).Value = WorksheetFunction.SumIfs(rng_1, rng_2, Cells(j, 1).Value, _
                rng_3, Cells(3, i).Value, _
                rng_5, SubBranch_Value)
        Next
        Counter_sum = Counter_sum + 1
    Next

ElseIf SubBranch_Value = "" And Branch_Value <> "" Then

    For i = 3 To LastCol_managers
        
        For j = 4 To LastRow
            Cells(j, i).Value = WorksheetFunction.SumIfs(rng_1, rng_2, Cells(j, 1).Value, _
                rng_3, Cells(3, i).Value, _
                rng_4, Branch_Value)
        Next
        Counter_sum = Counter_sum + 1
    Next

Else

    For i = 3 To LastCol_managers
        
        For j = 4 To LastRow
            Cells(j, i).Value = WorksheetFunction.SumIfs(rng_1, rng_2, Cells(j, 1).Value, _
                rng_3, Cells(3, i).Value, _
                rng_4, Branch_Value, _
                rng_5, SubBranch_Value)
        Next
        Counter_sum = Counter_sum + 1
    Next
    
End If

'***************************************************************

'копируем формулу, вставляем формат ячейки и оставляем только значения без формул
Cells(1, 2).Select
Selection.Copy
Range(Cells(4, 3), Cells(LastRow, LastCol_managers)).Select
Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False

Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False

'-------------------------------------------------------------------------------------------------
'Подсвечиваем 0 условных форматированием на всем диапазоне
Range(Cells(4, 3), Cells(LastRow, LastCol_managers)).Select
Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
    Formula1:="=0"
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
With Selection.FormatConditions(1).Font
    .Color = -16383844
    .TintAndShade = 0
End With
With Selection.FormatConditions(1).Interior
    .PatternColorIndex = xlAutomatic
    .Color = 13551615
    .TintAndShade = 0
End With
Selection.FormatConditions(1).StopIfTrue = False

'-------------------------------------------------------------------------------------------------
'Добавляем счет продаж по каждому филиалу

Sheets("1").Select
Number_of_stores = Cells(Rows.Count, 1).End(xlUp).Row

Counter = 1
Do While Counter <= Number_of_stores

    Sheets("1").Select
    Cells(Counter, 1).Select
    Selection.Copy
    Sheets("Tasks").Select
    LastCol = Cells(3, Columns.Count).End(xlToLeft).Column + 1
    Cells(3, LastCol).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Counter = Counter + 1
    
Loop

'-------------------------------------------------------------------------------------------------
'Считаем количество продающих менеджеров

LastCol = Cells(4, Columns.Count).End(xlToLeft).Column
Cells(4, LastCol).Select

'-----------------------------------------------------------------
'Переменная для определения границ сумм столбца Ранг
LastCol_for_sum = Cells(4, Columns.Count).End(xlToLeft).Column + 1
'-----------------------------------------------------------------

'реализация через цикл, так как не большой объем вычислений
Counter = 1
For i = 1 To Number_of_stores
    
    For j = 4 To LastRow
        
        Cells(j, i + LastCol).Value = WorksheetFunction.CountIfs(Range(Cells(j, 3), Cells(j, LastCol)), ">0", Range(Cells(2, 3), Cells(2, LastCol)), Cells(3, LastCol + Counter).Value)
    
    Next
    
    Counter = Counter + 1

Next

'-------------------------------------------------------------------------------------------------
'Добавляем столбец Ранг

LastCol = Cells(3, Columns.Count).End(xlToLeft).Column + 1
Cells(3, LastCol).Value = "Ранг"

For i = 4 To LastRow
        
    Cells(i, LastCol).Value = WorksheetFunction.Sum(Range(Cells(i, LastCol_for_sum), Cells(i, LastCol)))
    
Next

'-------------------------------------------------------------------------------------------------
'Изменяем форматирование вновь созданных столбцов

Cells(2, 2).Select
Selection.Copy
Range(Cells(3, LastCol_for_sum), Cells(3, LastCol)).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Cells(1, 2).Select
Selection.Copy
Range(Cells(4, LastCol_for_sum), Cells(LastRow, LastCol)).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Range(Cells(3, LastCol), Cells(LastRow, LastCol)).Select
Selection.Font.Bold = True
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent5
    .TintAndShade = 0.599993896298105
    .PatternTintAndShade = 0
End With

'-------------------------------------------------------------------------------------------------
'Копирование всей страницы на новый лист для удобства работы

    Range(Cells(2, 1), Cells(LastRow, LastCol)).Select
    Selection.Copy
    Sheets("Work").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With

    'автоматическая настройка ширины и внешнего вида отображения
    Range(Cells(2, LastCol_for_sum), Cells(LastRow, LastCol)).EntireColumn.Select
    Selection.Cut
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Application.CutCopyMode = False
        
    Range(Cells(2, 1), Cells(LastRow, LastCol)).EntireColumn.AutoFit
    ActiveWindow.Zoom = 70
    
    ColFilter = 0
    '----------------------------------------
    'Ищем столбец Ранг для будущей сортировки
    For i = 1 To LastCol
    
        If Cells(2, i).Value = "Ранг" Then
            ColFilter = i
            Exit For
        End If
        
    Next
    
    '----------------------------------------

    'добавляем фильтры и сортируем по столбцу Ранг по убыванию
    Rows("2:2").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Work").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Work").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        Cells(2, ColFilter), Cells(2, ColFilter)), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Work").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Cells(3, ColFilter + 1).Select
    ActiveWindow.FreezePanes = True
    
    '-------------------------------------------------------------------------------------------------    
    
    Cells(1, 1).Select
    Selection.Copy
    Cells(1, 2).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

End Sub