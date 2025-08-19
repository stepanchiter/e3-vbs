Option Explicit

' === Главная процедура === Запись атрибутов OOS из Excel
Sub WriteOOSAttributesFromExcel()
    Dim e3App, job, symbol
    Dim excelApp, excelWorkbook, excelSheet
    Dim filePath, i, rowNum, cellValue

    ' Переменные для хранения атрибутов
    Dim OOSTag, OOSType, OOSPras, OOSPnom, OOSInom
    Dim OOSDProizv3
    Dim OOSIras
    Dim OOSDProizv2 ' Переменная для атрибута ОД D_Proizv2
    Dim OOSDProizv1 ' Переменная для атрибута ОД D_Proizv1
    Dim OOSCos ' Переменная для атрибута ОД E_Cos

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Запись атрибутов OOS из Excel ==="

    ' 1. Запрос пути к файлу XLSX
    filePath = InputBox("Введите полный путь к файлу XLSX с данными:", "Путь к файлу Excel", "C:\MyData\OOSAttributes.xlsx")

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

    ' 3. Получение всех символов OOS из проекта и их сортировка
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

    ' Используем словарь для временного хранения OOS символов и их числовых индексов
    Dim OOSSymbolsMap
    Set OOSSymbolsMap = CreateObject("Scripting.Dictionary")
    Dim OOSNamesArray() ' Для хранения числовых индексов OOS, которые будем сортировать
    Dim OOSArrayCurrentSize : OOSArrayCurrentSize = 0 ' Текущее количество элементов в OOSNamesArray

    For i = LBound(allSymbolIds) To UBound(allSymbolIds) ' Используем LBound и UBound для надежности
        symbol.SetId(allSymbolIds(i))
        Dim symName : symName = symbol.GetName()
        If LCase(Left(symName, 3)) = "OOS" Then
            ' Извлекаем числовой индекс из имени OOS (например, из "OOS12" получаем 12)
            Dim OOSNum : OOSNum = CLng(Mid(symName, 4))
            If Not OOSSymbolsMap.Exists(CStr(OOSNum)) Then
                OOSSymbolsMap.Add CStr(OOSNum), allSymbolIds(i)
                
                ' Увеличиваем размер массива и переразмеряем его
                OOSArrayCurrentSize = OOSArrayCurrentSize + 1
                ReDim Preserve OOSNamesArray(OOSArrayCurrentSize - 1) ' Для 0-индексированного массива: N элементов -> максимальный индекс N-1
                
                OOSNamesArray(OOSArrayCurrentSize - 1) = OOSNum ' Присваиваем значение последнему элементу
            Else
                e3App.PutInfo 0, "ПРЕДУПРЕЖДЕНИЕ: Обнаружен дублирующийся OOS символ с номером '" & OOSNum & "'. Будет обработан только первый найденный."
            End If
        End If
    Next

    If OOSSymbolsMap.Count = 0 Then
        e3App.PutInfo 0, "В проекте не найдено символов OOS для записи атрибутов. Скрипт завершен."
        excelWorkbook.Close False
        excelApp.Quit
        Set excelSheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Set OOSSymbolsMap = Nothing
        Exit Sub
    End If

    ' Сортируем числовые индексы OOS символов по возрастанию
    Call SortNumericArrayAsc(OOSNamesArray) ' Используем вспомогательную функцию для сортировки

    e3App.PutInfo 0, "Найдено " & OOSSymbolsMap.Count & " символов OOS. Начинаем запись атрибутов..."

    ' 4. Цикл по отсортированным OOS символам и запись атрибутов
    For i = LBound(OOSNamesArray) To UBound(OOSNamesArray) ' Используем LBound и UBound для итерации по отсортированному массиву
        Dim currentOOSNum : currentOOSNum = OOSNamesArray(i)
        Dim currentOOSId : currentOOSId = OOSSymbolsMap.Item(CStr(currentOOSNum))

        symbol.SetId(currentOOSId)
        Dim currentSymName : currentSymName = symbol.GetName()

        ' Сопоставление OOS_N -> Строка N+1
        ' Если Excel начинается со строки 1, а OOS_N с 1, то строка N+1 - это 2, 3 и т.д.
        ' Если OOS_N это 0, то строка 0+1 = 1. Убедитесь, что это соответствует вашей структуре Excel.
        rowNum = currentOOSNum + 1 

        ' Чтение значений из Excel
        On Error Resume Next ' Включаем обработку ошибок для чтения ячеек
        OOSTag = Trim(CStr(excelSheet.Cells(rowNum, 13).Value))    ' Столбец M --> ОД E_TAG
        OOSType = Trim(CStr(excelSheet.Cells(rowNum, 14).Value))  ' Столбец N --> ОД E_TYPE
        OOSPras = Trim(CStr(excelSheet.Cells(rowNum, 6).Value)) ' Столбец F --> ОД E_Pras
        OOSPnom = Trim(CStr(excelSheet.Cells(rowNum, 5).Value))  ' Столбец E --> ОД E_Pnom
        OOSInom = Trim(CStr(excelSheet.Cells(rowNum, 8).Value)) ' Столбец H --> ОД E_Inom
        OOSDProizv3 = Trim(CStr(excelSheet.Cells(rowNum, 16).Value)) ' Столбец P --> ОД D_Proizv3
        OOSIras = Trim(CStr(excelSheet.Cells(rowNum, 9).Value))  ' Столбец I --> ОД E_Iras
        OOSDProizv2 = Trim(CStr(excelSheet.Cells(rowNum, 11).Value)) ' Столбец K --> ОД D_Proizv2
        OOSCos = Trim(CStr(excelSheet.Cells(rowNum, 4).Value))  ' Столбец D --> ОД E_Cos
        
        ' Замена "Преобразователь частоты" на "ПЧ" в тексте
        If InStr(1, OOSDProizv2, "Преобразователь частоты", vbTextCompare) > 0 Then
            OOSDProizv2 = Replace(OOSDProizv2, "Преобразователь частоты", "ПЧ", 1, -1, vbTextCompare)
        End If
        
        ' Если записывается ПЧ, то в D_Proizv1 записываем реле интерфейсное
        If InStr(1, OOSDProizv2, "ПЧ", vbTextCompare) > 0 Then
            OOSDProizv1 = "Реле интерфейсное 24 VDC, 1 CO с колодкой арт. RNC1CO024+SNB05-E-AR"
        Else
            OOSDProizv1 = "" ' Если не ПЧ, то оставляем пустым
        End If
        
        If Err.Number <> 0 Then
            e3App.PutInfo 0, "ПРЕДУПРЕЖДЕНИЕ: Ошибка при чтении данных для OOS символа '" & currentSymName & "' (ID: " & currentOOSId & ") из строки " & rowNum & ". Некоторые атрибуты могут быть пустыми. Ошибка: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' Запись атрибутов в символ OOS
        e3App.PutInfo 0, "  Обработка символа: " & currentSymName & " (ID: " & currentOOSId & ") -> Строка Excel: " & rowNum

        If Len(OOSTag) > 0 Then
            symbol.SetAttributeValue "ОД E_TAG", OOSTag
            e3App.PutInfo 0, "    Записано ОД E_TAG: " & OOSTag
        Else
            e3App.PutInfo 0, "    ОД E_TAG: <пусто>"
        End If

        If Len(OOSType) > 0 Then
            symbol.SetAttributeValue "ОД E_TYPE", OOSType
            e3App.PutInfo 0, "    Записано ОД E_TYPE: " & OOSType
        Else
            e3App.PutInfo 0, "    ОД E_TYPE: <пусто>"
        End If

        If Len(OOSPras) > 0 Then
            symbol.SetAttributeValue "ОД E_Pras", OOSPras
            e3App.PutInfo 0, "    Записано ОД E_Pras: " & OOSPras
        Else
            e3App.PutInfo 0, "    ОД E_Pras: <пусто>"
        End If

        If Len(OOSPnom) > 0 Then
            symbol.SetAttributeValue "ОД E_Pnom", OOSPnom
            e3App.PutInfo 0, "    Записано ОД E_Pnom: " & OOSPnom
        Else
            e3App.PutInfo 0, "    ОД E_Pnom: <пусто>"
        End If

        If Len(OOSInom) > 0 Then
            symbol.SetAttributeValue "ОД E_Inom", OOSInom
            e3App.PutInfo 0, "    Записано ОД E_Inom: " & OOSInom
        Else
            e3App.PutInfo 0, "    ОД E_Inom: <пусто>"
        End If

        If Len(OOSDProizv3) > 0 Then
            symbol.SetAttributeValue "ОД D_Proizv3", OOSDProizv3
            e3App.PutInfo 0, "    Записано ОД D_Proizv3: " & OOSDProizv3
        Else
            e3App.PutInfo 0, "    ОД D_Proizv3: <пусто>"
        End If

        If Len(OOSIras) > 0 Then
            symbol.SetAttributeValue "ОД E_Iras", OOSIras
            e3App.PutInfo 0, "    Записано ОД E_Iras: " & OOSIras
        Else
            e3App.PutInfo 0, "    ОД E_Iras: <пусто>"
        End If
        
        If Len(OOSDProizv2) > 0 Then
            symbol.SetAttributeValue "ОД D_Proizv2", OOSDProizv2
            e3App.PutInfo 0, "    Записано ОД D_Proizv2: " & OOSDProizv2
        Else
            e3App.PutInfo 0, "    ОД D_Proizv2: <пусто>"
        End If
        
        If Len(OOSDProizv1) > 0 Then
            symbol.SetAttributeValue "ОД D_Proizv1", OOSDProizv1
            e3App.PutInfo 0, "    Записано ОД D_Proizv1: " & OOSDProizv1
        Else
            e3App.PutInfo 0, "    ОД D_Proizv1: <пусто>"
        End If
        
        If Len(OOSCos) > 0 Then
            symbol.SetAttributeValue "ОД E_Cos", OOSCos
            e3App.PutInfo 0, "    Записано ОД E_Cos: " & OOSCos
        Else
            e3App.PutInfo 0, "    ОД E_Cos: <пусто>"
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
    Set OOSSymbolsMap = Nothing
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
Call WriteOOSAttributesFromExcel()