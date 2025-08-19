Option Explicit

' === Главная процедура === Переименование символов OOS по позиции
Sub RenameOOSSymbolsByLocation()
    Dim e3App, job, symbol, sheet
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName
    Dim s, i, j
    Dim OOSCounter : OOSCounter = 0 ' Счетчик для последовательного именования

    ' Переменные для хранения данных OOS символов для сортировки
    ' Каждый элемент в OOSSymbolsToRename будет массивом:
    ' (SymbolID, SheetID, SheetName, Column, Row, X, Y, OriginalName)
    Dim OOSSymbolsToRename()
    ReDim OOSSymbolsToRename(0) ' Инициализация с фиктивным элементом, будет скорректировано

    Dim OOSCountPlaced : OOSCountPlaced = 0 ' Счетчик только для размещенных OOS символов

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    Set sheet = job.CreateSheetObject() ' Для получения имени листа

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Переименование OOS символов по позиции ==="

    ' Получаем все ID символов в проекте
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "В проекте нет символов для анализа. Скрипт завершен."
        Set symbol = Nothing
        Set sheet = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    e3App.PutInfo 0, "Найдено " & allSymbolCount & " символов в проекте. Собираем данные и временно переименовываем OOS символы..."

    ' === Первый проход: Сбор данных о размещенных OOS символах и временное переименование ===
    For s = 1 To allSymbolCount
        currentSymbolId = allSymbolIds(s)
        symbol.SetId(currentSymbolId)
        symbolName = symbol.GetName()

        If LCase(Left(symbolName, 3)) = "oos" Then
            Dim xPos, yPos, gridDesc, columnValue, rowValue
            Dim sheetId : sheetId = symbol.GetSchemaLocation(xPos, yPos, gridDesc, columnValue, rowValue)

            If sheetId > 0 Then ' Символ размещен на схеме
                OOSCountPlaced = OOSCountPlaced + 1
                ReDim Preserve OOSSymbolsToRename(OOSCountPlaced) ' Увеличиваем размер массива

                sheet.SetId sheetId
                Dim sheetName : sheetName = sheet.GetName()

                ' Store the data for sorting, including original name
                OOSSymbolsToRename(OOSCountPlaced) = Array(currentSymbolId, sheetId, sheetName, columnValue, rowValue, xPos, yPos, symbolName)
                e3App.PutInfo 0, "  OOS символ найден на схеме: " & symbolName & " (ID: " & currentSymbolId & ") на листе: " & sheetName & " " & columnValue & rowValue & " (" & xPos & ", " & yPos & ")"
                
                ' Временно переименовываем символ, чтобы освободить имена OOS, OOS2 и т.д.
                ' Используем сокращенное временное имя, чтобы оно не превышало 12 символов
                symbol.SetName "OOSTMP" & Right(CStr(currentSymbolId), 6)
                e3App.PutInfo 0, "  Временно переименован в 'OOSTMP" & Right(CStr(currentSymbolId), 6) & "'."
            Else
                e3App.PutInfo 0, "  OOS символ '" & symbolName & "' (ID: " & currentSymbolId & ") НЕ размещен на схеме. Он не будет переименован этим скриптом."
            End If
        End If
    Next

    If OOSCountPlaced = 0 Then
        e3App.PutInfo 0, "Не найдено размещенных OOS символов на схеме для сортировки и переименования."
        e3App.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="
        Set symbol = Nothing
        Set sheet = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    e3App.PutInfo 0, "Найдено " & OOSCountPlaced & " размещенных OOS символов. Сортировка по позиции (от меньшей к большей)..."

    ' === Сортировка OOSSymbolsToRename (Bubble Sort) ===
    ' (SymbolID, SheetID, SheetName, Column, Row, X, Y, OriginalName)
    ' Index:   0       1         2       3      4   5   6        7
    For i = 1 To OOSCountPlaced - 1
        For j = i + 1 To OOSCountPlaced
            Dim item1Array, item2Array
            item1Array = OOSSymbolsToRename(i)
            item2Array = OOSSymbolsToRename(j)

            ' Comparison logic for ASCENDING order: SheetName (numerically), then Column (string), then Row (string), then X (numeric), then Y (numeric)
            ' If item1 should come AFTER item2 in ASCENDING order, we swap.
            
            ' Сравнение по имени листа (числовое)
            If CLng(item1Array(2)) > CLng(item2Array(2)) Then
                Call SwapArrayElements(OOSSymbolsToRename, i, j)
            ElseIf CLng(item1Array(2)) = CLng(item2Array(2)) Then
                ' Сравнение по столбцу (текстовое, без учета регистра)
                If StrComp(item1Array(3), item2Array(3), vbTextCompare) > 0 Then
                    Call SwapArrayElements(OOSSymbolsToRename, i, j)
                ElseIf StrComp(item1Array(3), item2Array(3), vbTextCompare) = 0 Then
                    ' Сравнение по строке (текстовое, без учета регистра)
                    If StrComp(item1Array(4), item2Array(4), vbTextCompare) > 0 Then
                        Call SwapArrayElements(OOSSymbolsToRename, i, j)
                    ElseIf StrComp(item1Array(4), item2Array(4), vbTextCompare) = 0 Then
                        ' Сравнение по X позиции (числовое)
                        If item1Array(5) > item2Array(5) Then
                            Call SwapArrayElements(OOSSymbolsToRename, i, j)
                        ElseIf item1Array(5) = item2Array(5) Then
                            If item1Array(6) > item2Array(6) Then ' Compare Y position (ascending)
                                Call SwapArrayElements(OOSSymbolsToRename, i, j)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    e3App.PutInfo 0, "Сортировка завершена. Приступаем к окончательному переименованию OOS символов."

    ' === Окончательное переименование OOS символов в отсортированном порядке ===
    For s = 1 To OOSCountPlaced
        OOSCounter = OOSCounter + 1
        currentSymbolId = OOSSymbolsToRename(s)(0) ' Get Symbol ID from sorted array
        Dim originalSymbolName : originalSymbolName = OOSSymbolsToRename(s)(7) ' Get original name for logging
        
        symbol.SetId(currentSymbolId)
        
        Dim newSymbolName : newSymbolName = "OOS" & OOSCounter
        
        ' Переименовываем символ
        symbol.SetName newSymbolName
        e3App.PutInfo 0, "  Символ '" & originalSymbolName & "' (ID: " & currentSymbolId & ") переименован в '" & newSymbolName & "' (на листе: " & OOSSymbolsToRename(s)(2) & " " & OOSSymbolsToRename(s)(3) & OOSSymbolsToRename(s)(4) & ")."
    Next

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА: Переименовано " & OOSCounter & " OOS символов ==="

    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' Helper Sub: Swaps two elements in an array
Sub SwapArrayElements(arr, index1, index2)
    Dim temp
    temp = arr(index1)
    arr(index1) = arr(index2)
    arr(index2) = temp
End Sub

' === Основной запуск ===
Call RenameOOSSymbolsByLocation()
