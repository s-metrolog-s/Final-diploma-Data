Function CheckInputFiles() As Boolean
'
' Проверка на наличие всех необходимых файлов для работы
'
'----------------------------------------------------------
' По тех. заданию заложить как минимум +1 файл для проверки
' Дублируем первый файл как заглушку для будущего добавления
'----------------------------------------------------------
'

Dim File1Path As String
Dim File2Path As String

File1Path = ThisWorkbook.Path & "/data.xlsx"
File2Path = ThisWorkbook.Path & "/data.xlsx"

If Dir(File1Path, vbDirectory) = vbNullString Or _
    Dir(File2Path, vbDirectory) = vbNullString Then
    
    MsgBox "Необходимые файлы для запуска отсуствуют" & vbNewLine & "Проверьте наличие файлов" & vbNewLine & _
        "data.xlsx" & vbNewLine & _
        "data.xlsx"
        
    CheckInputFiles = False
    Exit Function

End If

    CheckInputFiles = True

End Function