Sub SecondStepWithData()
'
' Модуль запуска для старта формирования результирующей таблицы и дополнительных расчетов в потенциале
'

'выключаем обновление экрана
Application.ScreenUpdating = False
Application.CutCopyMode = False
Sheets("1").Visible = True
Sheets("2").Visible = True
Sheets("3").Visible = True
Sheets("4").Visible = True
Sheets("5").Visible = True
Sheets("data").Visible = True
Sheets("Tasks").Visible = True
    
Dim t As Single
t = Timer
            
    MakeTable ' Заполнение результирующей таблицы

    Sheets("Settings").Select
    
    t = Timer - t
    UpdateMsg = MsgBox("Обновление прошло успешно за " & Round(t, 0) & " секунд", vbOKOnly, "Обновление")
    
Sheets("1").Visible = False
Sheets("2").Visible = False
Sheets("3").Visible = False
Sheets("4").Visible = False
Sheets("5").Visible = False
Sheets("data").Visible = False
Sheets("Tasks").Visible = False
Application.ScreenUpdating = True
    
End Sub