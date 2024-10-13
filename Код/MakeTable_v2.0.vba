Sub MakeTable()
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
'!новая реалиазация через вставку готовой формулы
'!время выполнения уменьшено в 50 раз!
'!текущий блок сокращен на 70 строк

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Sheets("Tasks").Select

' Использование глобальных переменных отменено по ТЗ, реализована возможность настройки
' пользователем номеров столбцов во входном файле, так как у части учетных записей
' данные могут выгружаться в другом порядке, загружаем с листа Settings

Sum_column = Sheets("Settings").Cells(5, 6).Value
Article_column = Sheets("Settings").Cells(4, 6).Value
Manager_column = Sheets("Settings").Cells(3, 6).Value
Brunch_column = Sheets("Settings").Cells(6, 6).Value
SubBrunch_column = Sheets("Settings").Cells(7, 6).Value

'копируем значения сегментов и подсегментов для фильтрации данных
Branch_Value = Sheets("Settings").Range("J2").Value
SubBranch_Value = Sheets("Settings").Range("J5").Value

'для удобства внесения формулы меняем нотацию
Application.ReferenceStyle = xlR1C1

If Branch_Value = "" And SubBranch_Value = "" Then

    Range("B1").FormulaR1C1Local = "=СУММЕСЛИМН(data!C" & Sum_column & _
    ";data!C" & Article_column & ";Tasks!RC1;data!C" & Manager_column & ";Tasks!R3C)"
    
ElseIf Branch_Value = "" And SubBranch_Value <> "" Then

    Range("B1").FormulaR1C1Local = "=СУММЕСЛИМН(data!C" & Sum_column & _
    ";data!C" & Article_column & ";Tasks!RC1;data!C" & Manager_column & _
    ";Tasks!R3C;data!C" & SubBrunch_column & ";Settings!R5C10)"

ElseIf SubBranch_Value = "" And Branch_Value <> "" Then

    Range("B1").FormulaR1C1Local = "=СУММЕСЛИМН(data!C" & Sum_column & _
    ";data!C" & Article_column & ";Tasks!RC1;data!C" & Manager_column & _
    ";Tasks!R3C;data!C" & Brunch_column & ";Settings!R2C10)"

Else

    Range("B1").FormulaR1C1Local = "=СУММЕСЛИМН(data!C" & Sum_column & _
    ";data!C" & Article_column & ";Tasks!RC1;data!C" & Manager_column & _
    ";Tasks!R3C;data!C" & Brunch_column & ";Settings!R2C10;data!C" & SubBrunch_column & ";Settings!R5C10)"
        
End If

'возращаем нотацию по умолчанию
Application.ReferenceStyle = xlA1

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
    'Добавление столбца Сумма

    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight
    Cells(2, 3).Value = "Сумма"
    Cells(2, 2).Select
    Selection.Copy
    Cells(2, 3).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    LastCol_Total = Cells(3, Columns.Count).End(xlToLeft).Column
    LastRow_Total = Cells(Rows.Count, 1).End(xlUp).Row
    
    Counter = 1
    For i = 3 To LastRow_Total
    
        Cells(i, 3).Value = WorksheetFunction.Sum(Range(Cells(i, ColFilter + 2), Cells(i, LastCol_Total)))
        Counter = Counter + 1

    Next
    
    Columns("C:C").Select
    Selection.NumberFormat = "#,##0"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("C:C").EntireColumn.AutoFit

    '-------------------------------------------------------------------------------------------------
    'Добавление столбца Маржа
    'определяем номер столбца для нотации R1C1
    '(номер столбца на листе data) - (номер колонки, в которую записываем значения)
    
    Gross_column = (Sum_column + 1) - 4 '4 остается неизменным
    Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    Cells(2, 4).Value = "Маржа"
    Cells(2, 2).Select
    Selection.Copy
    Cells(2, 4).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    'Считаем маржу по каждому артикулу
    Application.ReferenceStyle = xlR1C1
    Counter = 1
    For i = 3 To LastRow_Total
        
        Cells(i, 4).FormulaR1C1Local = "=СУММЕСЛИМН(data!C[" & Gross_column & "];data!C;Work!RC[-3])"
        
        Counter = Counter + 1

    Next
    Application.ReferenceStyle = xlA1
    
    'форматируем данные к требуемому виду
    Columns("D:D").Select
    Selection.NumberFormat = "#,##0"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("D:D").EntireColumn.AutoFit
    
    'снимаем формулы
    Range(Cells(3, 4), Cells(LastRow_Total, 4)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    '-------------------------------------------------------------------------------------------------
    
    Cells(1, 1).Select
    Selection.Copy
    Cells(1, 2).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False

End Sub