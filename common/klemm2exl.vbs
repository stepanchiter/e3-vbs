'==============================================
' Генерация списка клемм XT, вывод и экспорт в Excel
' Ред. [текущая дата]
'==============================================

' Основной код
Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject
Set Dev = Job.CreateDeviceObject

' Указываем путь к Excel-файлу (измените на ваш реальный путь)
'Dim excelFilePath
'excelFilePath = "\\Vt5\niokr\ZAKAZI2016\ZN_1108066\brk66976_241073206.xls"  ' Укажите правильный путь здесь!

' Формируем имя файла автоматически
Dim excelFilePath
excelFilePath = GetExcelFileName(Job)
If excelFilePath = "" Then
    WScript.Quit
End If

' Получаем ID всех клеммников
terminalCount = Job.GetTerminalIds(DevIds)

If terminalCount = 0 Then
    App.PutInfo 1, "Клеммники не найдены!"
    WScript.Quit
End If

' Обработка данных
Dim xtTerminals
xtTerminals = ProcessTerminals(Job, Dev, DevIds, terminalCount)

' Проверяем, что получили массив перед выводом и экспортом
If IsArray(xtTerminals) Then
    ' Выводим массив для проверки
    Call ShowArrayBeforeExport(xtTerminals)
    
    ' Экспорт в Excel с проверкой
    Call ExportToExcel(xtTerminals, excelFilePath, "Маркировки", "E2")
Else
    App.PutInfo 1, "Нет данных для экспорта (XT-клеммники не найдены)"
End If

' Завершение
Set Dev = Nothing
Set Job = Nothing
Set App = Nothing
WScript.Quit

'==============================================
' Процедура обработки клеммников
' Возвращает массив XT клеммников или Nothing
'==============================================
Function ProcessTerminals(Job, Dev, DevIds, totalCount)
    ' 1. Сбор данных
    Dim terminals(), xtTerminals()
    Dim i, j, xtCount
    
    ReDim terminals(totalCount - 1, 1)
    xtCount = 0
    
    For i = 0 To totalCount - 1
        Dev.SetId DevIds(i + 1)
        terminals(i, 0) = Dev.GetMasterPinName
        terminals(i, 1) = Dev.GetName
        
        If InStr(1, UCase(terminals(i, 1)), "XT", vbTextCompare) > 0 Then
            xtCount = xtCount + 1
        End If
    Next
    
    If xtCount = 0 Then
        ProcessTerminals = Null ' Возвращаем Null вместо Nothing
        Exit Function
    End If
    
    ' 2. Фильтрация XT клеммников
    ReDim xtTerminals(xtCount - 1, 1)
    j = 0
    
    For i = 0 To totalCount - 1
        If InStr(1, UCase(terminals(i, 1)), "XT", vbTextCompare) > 0 Then
            xtTerminals(j, 0) = terminals(i, 0)
            xtTerminals(j, 1) = terminals(i, 1)
            j = j + 1
        End If
    Next
    
    ' 3. Сортировка
    Dim temp1, temp2
    For i = UBound(xtTerminals, 1) - 1 To 0 Step -1
        For j = 0 To i
            If StrComp(xtTerminals(j, 1), xtTerminals(j + 1, 1), vbTextCompare) > 0 Then
                temp1 = xtTerminals(j, 0)
                temp2 = xtTerminals(j, 1)
                xtTerminals(j, 0) = xtTerminals(j + 1, 0)
                xtTerminals(j, 1) = xtTerminals(j + 1, 1)
                xtTerminals(j + 1, 0) = temp1
                xtTerminals(j + 1, 1) = temp2
            End If
        Next
    Next
    
    ProcessTerminals = xtTerminals
End Function


'==============================================
' Процедура вывода массива перед экспортом
'==============================================
Sub ShowArrayBeforeExport(dataArray)
    Dim i, outputStr
    outputStr = "Массив для экспорта (" & UBound(dataArray, 1) + 1 & " строк):" & vbCrLf
    outputStr = outputStr & "=========================" & vbCrLf
    outputStr = outputStr & "Индекс | Номер клеммы | Поз. обозначение" & vbCrLf
    outputStr = outputStr & "-------------------------" & vbCrLf
    
    For i = 0 To UBound(dataArray, 1)
        outputStr = outputStr & Right("   " & i, 4) & " | " & _
                   Right("      " & dataArray(i, 0), 11) & " | " & _
                   dataArray(i, 1) & vbCrLf
    Next
    
    outputStr = outputStr & "========================="
    App.PutInfo 0, outputStr
End Sub

'==============================================
' Процедура экспорта в Excel (с очисткой и заливкой)
'==============================================
Sub ExportToExcel(dataArray, filePath, sheetName, startCell)
    On Error Resume Next
    
    Dim ExcelApp, ExcelBook, ExcelSheet
    Dim i, colNum, rowNum, lastRow
    
    ' Разбираем адрес ячейки
    colNum = Asc(UCase(Mid(startCell, 1, 1))) - 64 ' A=1, B=2...
    rowNum = CInt(Mid(startCell, 2))
    
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = True ' Для отладки
    
    ' Пытаемся открыть файл по указанному пути
    App.PutInfo 0, "Пытаюсь открыть файл: " & filePath
    Set ExcelBook = ExcelApp.Workbooks.Open(filePath)
    
    If Err.Number <> 0 Then
        App.PutInfo 1, "Ошибка открытия файла " & filePath & ": " & Err.Description
        ExcelApp.Quit
        Set ExcelApp = Nothing
        WScript.Quit
    End If
    On Error GoTo 0
    
    ' Ищем нужный лист
    Set ExcelSheet = Nothing
    On Error Resume Next
    Set ExcelSheet = ExcelBook.Sheets(sheetName)
    If Err.Number <> 0 Then
        App.PutInfo 1, "Лист '" & sheetName & "' не найден в файле " & filePath
        ExcelBook.Close False
        ExcelApp.Quit
        Set ExcelApp = Nothing
        WScript.Quit
    End If
    On Error GoTo 0
    
    ' Очищаем столбец E начиная с E2
    App.PutInfo 0, "Очищаю столбец E перед вставкой данных..."
    With ExcelSheet
        ' Находим последнюю заполненную строку в столбце E
        lastRow = .Cells(.Rows.Count, 5).End(-4162).Row ' -4162 = xlUp
        
        ' Если есть данные ниже E2 - очищаем
        If lastRow >= rowNum Then
            .Range(.Cells(rowNum, 5), .Cells(lastRow, 5)).ClearContents
        End If
    End With
    
    ' Записываем данные
    App.PutInfo 0, "Записываю " & UBound(dataArray, 1) + 1 & " строк в столбец E..."
    For i = 0 To UBound(dataArray, 1)
        ExcelSheet.Cells(rowNum + i, colNum).Value = dataArray(i, 0)
    Next
    
    ' Закрашиваем столбец E в желтый цвет (от E2 до последней заполненной ячейки)
    App.PutInfo 0, "Закрашиваю столбец E в желтый цвет..."
    With ExcelSheet
        lastRow = .Cells(.Rows.Count, 5).End(-4162).Row
        If lastRow < rowNum Then lastRow = rowNum ' Если данных нет, закрашиваем только E2
        
        With .Range(.Cells(rowNum, 5), .Cells(lastRow, 5)).Interior
            .Color = 65535 ' Желтый цвет
            .Pattern = 1   ' Сплошная заливка
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
    
    ' Сохраняем и закрываем
    ExcelBook.Save
    ExcelBook.Close
    ExcelApp.Quit
    
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    
    App.PutInfo 0, "Экспорт завершен: данные записаны и столбец E закрашен в " & filePath
End Sub


'==============================================
' Процедура формирования имени XLS файла
' Возвращает полный путь к файлу в формате:
' <путь_проекта>\<имя_проекта>.xls
' с заменой "Sch2_" на "brk" и удалением ".e3d"
'==============================================
Function GetExcelFileName(Job)
    Dim projectPath, projectName, excelFileName
    
    ' Получаем путь проекта
    projectPath = Job.GetPath()
    If Len("" & projectPath) = 0 Then
        App.PutInfo 1, "Ошибка получения пути проекта"
        GetExcelFileName = ""
        Exit Function
    End If
    
    ' Получаем имя проекта
    projectName = Job.GetName()
    If Len("" & projectName) = 0 Then
        App.PutInfo 1, "Ошибка получения имени проекта"
        GetExcelFileName = ""
        Exit Function
    End If
    
    ' Заменяем "Sch2_" на "brk" в имени проекта
    projectName = Replace(projectName, "Sch2_", "brk", 1, -1, vbTextCompare)
    
    ' Убираем расширение .e3d если есть (регистронезависимо)
    projectName = Replace(projectName, ".e3d", "", 1, -1, vbTextCompare)
    projectName = Replace(projectName, ".E3D", "", 1, -1, vbTextCompare)
    
    ' Создаем имя файла Excel (без лишних суффиксов)
    excelFileName = projectPath & "\" & projectName & ".xls"
    
    App.PutInfo 0, "Сформирован путь к файлу Excel: " & excelFileName
    GetExcelFileName = excelFileName
End Function