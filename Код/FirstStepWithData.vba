Sub FirstStepWithData()
'
' Запуск всех модулей, связаннных с обновлением входных данных
'
    Dim checkFiles As Boolean
    
    'проверяем наличие файла с исходными данными в папке с макросом
    checkFiles = CheckInputFiles()
    
    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    
    Sheets("1").Visible = True
    Sheets("2").Visible = True
    Sheets("3").Visible = True
    Sheets("4").Visible = True
    Sheets("5").Visible = True
    Sheets("data").Visible = True
    Sheets("Tasks").Visible = True
    
    'обновляем только в случае наличия входных данных
    If checkFiles Then
    
        Dim t As Single
        t = Timer
         
        ClearAllData ' Очистка всех листов для загрузки данных
        UpdateInputData ' Обновляем файлы с продажами и маржой
        FindUniqueInData ' Отбор уникальных элементов для построения итоговой таблицы
        UpdateBranchFilters ' Обновление фильтрации
        Sheets("Settings").Select
    
        t = Timer - t
        UpdateMsg = MsgBox("Обновление прошло успешно за " & Round(t, 0) & " секунд", vbOKOnly, "Обновление")
    
    End If
    
    Sheets("1").Visible = False
    Sheets("2").Visible = False
    Sheets("3").Visible = False
    Sheets("4").Visible = False
    Sheets("5").Visible = False
    Sheets("data").Visible = False
    Sheets("Tasks").Visible = False
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

End Sub