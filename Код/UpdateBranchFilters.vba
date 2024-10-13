Sub UpdateBranchFilters()
'
' Проверка значений для сегментов и подсегментов
' создание выпадающих списков для фильтрации
'
    'сегменты
    Sheets("4").Select
    LastRow_brunch = Cells(Rows.Count, 1).End(xlUp).Row
    Sheets("Settings").Select
    Range("J2:J4").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='4'!$A$1:$A$" & LastRow_brunch
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    'подсегменты
    Sheets("5").Select
    LastRow_subbrunch = Cells(Rows.Count, 1).End(xlUp).Row
    Sheets("Settings").Select
    Range("J5:J7").Select
    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='5'!$A$1:$A$" & LastRow_subbrunch
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
End Sub