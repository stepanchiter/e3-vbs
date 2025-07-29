Option Explicit

' === Главная процедура === Запись атрибутов OOO из Excel
Sub WriteOOOAttributesFromExcel()
    Dim e3App, job, symbol
    Dim excelApp, excelWorkbook, excelSheet
    Dim filePath, i, rowNum, cellValue

    ' Переменные для хранения атрибутов
    Dim oooTag, oooType, oooPras, oooPnom, oooInom
    Dim oooDProizv3
    Dim oooIras
    Dim oooDProizv2 ' Переменная для атрибута ОД D_Proizv2
    Dim oooDProizv1 ' Переменная для атрибута ОД D_Proizv1

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Запись атрибутов OOO из Excel ==="

    ' 1. Запрос пути к файлу XLSX
    filePath = InputBox("Введите полный путь к файлу XLSX с данными:", "Путь к файлу Excel", "C:\MyData\OooAttributes.xlsx")

    If Trim(filePath) = "" Then
        e3App.PutInfo 0, "Путь к файлу не был введен. Скрипт отменен."
        Exit Sub
    End If

    ' Проверка существования файла
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        e3App.PutInfo 0, "ОШИБКА: Файл '" & filePath & "' не найден. Проверьте путь."
        Set fso = Nothing
        Exit Sub
    End If
    Set fso = Nothing

    ' 2. Запуск Excel и открытие файла
    On Error Resume Next ' Включаем обработку ошибок для операций с Excel
    Set excelApp = GetObject("Excel.Application")
    If Err.Number <> 0 Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0 ' Выключаем обработку ошибок

    If excelApp Is Nothing Then
        e3App.PutInfo 0, "ОШИБКА: Не удалось запустить или подключиться к Excel. Убедитесь, что Excel установлен."
        Exit Sub
    End If

    excelApp.Visible = False ' Скрываем Excel для фоновой работы
    excelApp.DisplayAlerts = False ' Отключаем предупреждения (например, о режиме совместимости)

    On Error Resume Next
    Set excelWorkbook = excelApp.Workbooks.Open(filePath)
    If Err.Number <> 0 Then
        e3App.PutInfo 0, "ОШИБКА: Не удалось открыть файл Excel: '" & filePath & "'. Ошибка: " & Err.Description
        excelApp.Quit
        Set excelApp = Nothing
        Exit Sub
    End If
    On Error GoTo 0

    Set excelSheet = excelWorkbook.Sheets(1) ' Работаем с первым листом

    ' 3. Получение всех символов OOO из проекта и их сортировка
    Dim allSymbolIds()
    Dim allSymbolCount
    ' job.GetSymbolIds возвращает количество элементов и заполняет массив allSymbolIds
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "В проекте нет символов для анализа. Скрипт завершен."
        excelWorkbook.Close False ' Закрыть без сохранения
        excelApp.Quit
        Set excelSheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    ' Используем словарь для временного хранения OOO символов и их числовых индексов
    Dim oooSymbolsMap
    Set oooSymbolsMap = CreateObject("Scripting.Dictionary")
    Dim oooNamesArray() ' Для хранения числовых индексов OOO, которые будем сортировать
    Dim oooArrayCurrentSize : oooArrayCurrentSize = 0 ' Текущее количество элементов в oooNamesArray

    For i = LBound(allSymbolIds) To UBound(allSymbolIds) ' Используем LBound и UBound для надежности
        symbol.SetId(allSymbolIds(i))
        Dim symName : symName = symbol.GetName()
        If LCase(Left(symName, 3)) = "ooo" Then
            ' Извлекаем числовой индекс из имени OOO (например, из "OOO12" получаем 12)
            Dim oooNum : oooNum = CLng(Mid(symName, 4))
            If Not oooSymbolsMap.Exists(CStr(oooNum)) Then
                oooSymbolsMap.Add CStr(oooNum), allSymbolIds(i)
                
                ' Увеличиваем размер массива и переразмеряем его
                oooArrayCurrentSize = oooArrayCurrentSize + 1
                ReDim Preserve oooNamesArray(oooArrayCurrentSize - 1) ' Для 0-индексированного массива: N элементов -> максимальный индекс N-1
                
                oooNamesArray(oooArrayCurrentSize - 1) = oooNum ' Присваиваем значение последнему элементу
            Else
                e3App.PutInfo 0, "ПРЕДУПРЕЖДЕНИЕ: Обнаружен дублирующийся OOO символ с номером '" & oooNum & "'. Будет обработан только первый найденный."
            End If
        End If
    Next

    If oooSymbolsMap.Count = 0 Then
        e3App.PutInfo 0, "В проекте не найдено символов OOO для записи атрибутов. Скрипт завершен."
        excelWorkbook.Close False
        excelApp.Quit
        Set excelSheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Set oooSymbolsMap = Nothing
        Exit Sub
    End If

    ' Сортируем числовые индексы OOO символов по возрастанию
    Call SortNumericArrayAsc(oooNamesArray) ' Используем вспомогательную функцию для сортировки

    e3App.PutInfo 0, "Найдено " & oooSymbolsMap.Count & " символов OOO. Начинаем запись атрибутов..."

    ' 4. Цикл по отсортированным OOO символам и запись атрибутов
    For i = LBound(oooNamesArray) To UBound(oooNamesArray) ' Используем LBound и UBound для итерации по отсортированному массиву
        Dim currentOOONum : currentOOONum = oooNamesArray(i)
        Dim currentOOOId : currentOOOId = oooSymbolsMap.Item(CStr(currentOOONum))

        symbol.SetId(currentOOOId)
        Dim currentSymName : currentSymName = symbol.GetName()

        ' Сопоставление OOO_N -> Строка N+1
        ' Если Excel начинается со строки 1, а OOO_N с 1, то строка N+1 - это 2, 3 и т.д.
        ' Если OOO_N это 0, то строка 0+1 = 1. Убедитесь, что это соответствует вашей структуре Excel.
        rowNum = currentOOONum + 1 

        ' Чтение значений из Excel
        On Error Resume Next ' Включаем обработку ошибок для чтения ячеек
        oooTag = Trim(CStr(excelSheet.Cells(rowNum, 13).Value))    ' Столбец M --> ОД E_TAG
        oooType = Trim(CStr(excelSheet.Cells(rowNum, 14).Value))  ' Столбец N --> ОД E_TYPE
        oooPras = Trim(CStr(excelSheet.Cells(rowNum, 6).Value)) ' Столбец F --> ОД E_Pras
        oooPnom = Trim(CStr(excelSheet.Cells(rowNum, 5).Value))  ' Столбец E --> ОД E_Pnom
        oooInom = Trim(CStr(excelSheet.Cells(rowNum, 8).Value)) ' Столбец H --> ОД E_Inom
        oooDProizv3 = Trim(CStr(excelSheet.Cells(rowNum, 16).Value)) ' Столбец P --> ОД D_Proizv3
        oooIras = Trim(CStr(excelSheet.Cells(rowNum, 9).Value))  ' Столбец I --> ОД E_Iras
        oooDProizv2 = Trim(CStr(excelSheet.Cells(rowNum, 11).Value)) ' Столбец K --> ОД D_Proizv2
        
        ' Замена "Преобразователь частоты" на "ПЧ" в тексте
        If InStr(1, oooDProizv2, "Преобразователь частоты", vbTextCompare) > 0 Then
            oooDProizv2 = Replace(oooDProizv2, "Преобразователь частоты", "ПЧ", 1, -1, vbTextCompare)
        End If
        
        ' Если записывается ПЧ, то в D_Proizv1 записываем реле интерфейсное
        If InStr(1, oooDProizv2, "ПЧ", vbTextCompare) > 0 Then
            oooDProizv1 = "Реле интерфейсное 24 VDC, 1 CO с колодкой арт. RNC1CO024+SNB05-E-AR"
        Else
            oooDProizv1 = "" ' Если не ПЧ, то оставляем пустым
        End If
        
        If Err.Number <> 0 Then
            e3App.PutInfo 0, "ПРЕДУПРЕЖДЕНИЕ: Ошибка при чтении данных для OOO символа '" & currentSymName & "' (ID: " & currentOOOId & ") из строки " & rowNum & ". Некоторые атрибуты могут быть пустыми. Ошибка: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' Запись атрибутов в символ OOO
        e3App.PutInfo 0, "  Обработка символа: " & currentSymName & " (ID: " & currentOOOId & ") -> Строка Excel: " & rowNum

        If Len(oooTag) > 0 Then
            symbol.SetAttributeValue "ОД E_TAG", oooTag
            e3App.PutInfo 0, "    Записано ОД E_TAG: " & oooTag
        Else
            e3App.PutInfo 0, "    ОД E_TAG: <пусто>"
        End If

        If Len(oooType) > 0 Then
            symbol.SetAttributeValue "ОД E_TYPE", oooType
            e3App.PutInfo 0, "    Записано ОД E_TYPE: " & oooType
        Else
            e3App.PutInfo 0, "    ОД E_TYPE: <пусто>"
        End If

        If Len(oooPras) > 0 Then
            symbol.SetAttributeValue "ОД E_Pras", oooPras
            e3App.PutInfo 0, "    Записано ОД E_Pras: " & oooPras
        Else
            e3App.PutInfo 0, "    ОД E_Pras: <пусто>"
        End If

        If Len(oooPnom) > 0 Then
            symbol.SetAttributeValue "ОД E_Pnom", oooPnom
            e3App.PutInfo 0, "    Записано ОД E_Pnom: " & oooPnom
        Else
            e3App.PutInfo 0, "    ОД E_Pnom: <пусто>"
        End If

        If Len(oooInom) > 0 Then
            symbol.SetAttributeValue "ОД E_Inom", oooInom
            e3App.PutInfo 0, "    Записано ОД E_Inom: " & oooInom
        Else
            e3App.PutInfo 0, "    ОД E_Inom: <пусто>"
        End If

        If Len(oooDProizv3) > 0 Then
            symbol.SetAttributeValue "ОД D_Proizv3", oooDProizv3
            e3App.PutInfo 0, "    Записано ОД D_Proizv3: " & oooDProizv3
        Else
            e3App.PutInfo 0, "    ОД D_Proizv3: <пусто>"
        End If

        If Len(oooIras) > 0 Then
            symbol.SetAttributeValue "ОД E_Iras", oooIras
            e3App.PutInfo 0, "    Записано ОД E_Iras: " & oooIras
        Else
            e3App.PutInfo 0, "    ОД E_Iras: <пусто>"
        End If
        
        If Len(oooDProizv2) > 0 Then
            symbol.SetAttributeValue "ОД D_Proizv2", oooDProizv2
            e3App.PutInfo 0, "    Записано ОД D_Proizv2: " & oooDProizv2
        Else
            e3App.PutInfo 0, "    ОД D_Proizv2: <пусто>"
        End If
        
        If Len(oooDProizv1) > 0 Then
            symbol.SetAttributeValue "ОД D_Proizv1", oooDProizv1
            e3App.PutInfo 0, "    Записано ОД D_Proizv1: " & oooDProizv1
        Else
            e3App.PutInfo 0, "    ОД D_Proizv1: <пусто>"
        End If
    Next

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА: Атрибуты успешно записаны ==="

    ' 5. Очистка объектов Excel
    excelWorkbook.Close False ' Закрыть без сохранения
    excelApp.Quit
    
    Set excelSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
    Set oooSymbolsMap = Nothing
End Sub

' === Вспомогательная процедура === Сортировка числового массива по возрастанию
Sub SortNumericArrayAsc(arr)
    Dim i, j, temp
    ' Если массив пуст или содержит только 1 элемент, сортировка не нужна
    ' LBound(arr) и UBound(arr) корректно обрабатывают массив любой индексации
    If UBound(arr) < LBound(arr) + 1 Then Exit Sub

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next
    Next
End Sub

' === Основной запуск скрипта ===
Call WriteOOOAttributesFromExcel()