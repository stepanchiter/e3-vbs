'*******************************************************************************
' Название скрипта: E3_UZ_RenamePOOSymbolsByLocation_SortedSheets
' Автор: E3.series VBScript Assistant
' Дата: 01.07.2025
' Описание: Скрипт для автоматического переименования символов POO
'          на основе их расположения на листах, с корректной сортировкой листов
'          по числовому значению.
'*******************************************************************************
Option Explicit

' === Глобальные переменные ===
' Объявляем переменные глобально, чтобы они были доступны всем процедурам
Dim e3App, job, symbol, sheet

' === Главная процедура === Переименование символов POO по позиции
Sub RenamePOOSymbolsByLocation()
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName
    Dim s, i, j
    Dim pooCounter : pooCounter = 0 ' Счетчик для последовательного именования

    ' Переменные для хранения данных POO символов для сортировки
    ' Каждый элемент будет массивом:
    ' (SymbolID, SheetID, SheetName, Column, Row, X, Y, OriginalName)
    Dim pooSymbolsToRename()
    ReDim pooSymbolsToRename(0) ' Инициализация

    Dim pooCountPlaced : pooCountPlaced = 0 ' Счетчик только для размещенных POO символов

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    Set sheet = job.CreateSheetObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Переименование POO символов по позиции ==="

    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "В проекте нет символов для анализа. Скрипт завершен."
        Cleanup
        Exit Sub
    End If

    e3App.PutInfo 0, "Найдено " & allSymbolCount & " символов. Собираем данные и временно переименовываем POO символы..."

    ' === Сбор данных и временное переименование ===
    For s = 1 To allSymbolCount
        currentSymbolId = allSymbolIds(s)
        symbol.SetId(currentSymbolId)
        symbolName = symbol.GetName()

        If LCase(Left(symbolName, 3)) = "poo" Then
            Dim xPos, yPos, gridDesc, columnValue, rowValue
            Dim sheetId : sheetId = symbol.GetSchemaLocation(xPos, yPos, gridDesc, columnValue, rowValue)

            If sheetId > 0 Then
                pooCountPlaced = pooCountPlaced + 1
                ReDim Preserve pooSymbolsToRename(pooCountPlaced)

                sheet.SetId sheetId
                Dim sheetName : sheetName = sheet.GetName()

                pooSymbolsToRename(pooCountPlaced) = Array(currentSymbolId, sheetId, sheetName, columnValue, rowValue, xPos, yPos, symbolName)
                symbol.SetName "POOTMP" & Right(CStr(currentSymbolId), 6)
            End If
        End If
    Next

    If pooCountPlaced = 0 Then
        e3App.PutInfo 0, "Размещённые символы POO не найдены. Скрипт завершён."
        Cleanup
        Exit Sub
    End If

    e3App.PutInfo 0, "Найдено " & pooCountPlaced & " размещенных POO символов. Сортировка..."

    ' === Сортировка по SheetNumber (числовое), Column, Row, X, Y ===
    For i = 1 To pooCountPlaced - 1
        For j = i + 1 To pooCountPlaced
            Dim a1, a2
            a1 = pooSymbolsToRename(i)
            a2 = pooSymbolsToRename(j)

            ' Извлекаем числовые части названий листов для корректной сортировки
            Dim sheetNum1 : sheetNum1 = ExtractSheetNumber(a1(2))
            Dim sheetNum2 : sheetNum2 = ExtractSheetNumber(a2(2))

            If sheetNum1 > sheetNum2 Then
                SwapArrayElements pooSymbolsToRename, i, j
            ElseIf sheetNum1 = sheetNum2 Then
                If StrComp(a1(3), a2(3), vbTextCompare) > 0 Then ' Сравнение столбцов (Column)
                    SwapArrayElements pooSymbolsToRename, i, j
                ElseIf StrComp(a1(3), a2(3), vbTextCompare) = 0 Then
                    If StrComp(a1(4), a2(4), vbTextCompare) > 0 Then ' Сравнение строк (Row)
                        SwapArrayElements pooSymbolsToRename, i, j
                    ElseIf StrComp(a1(4), a2(4), vbTextCompare) = 0 Then
                        If a1(5) > a2(5) Then ' Сравнение X-координат
                            SwapArrayElements pooSymbolsToRename, i, j
                        ElseIf a1(5) = a2(5) And a1(6) > a2(6) Then ' Сравнение Y-координат
                            SwapArrayElements pooSymbolsToRename, i, j
                        End If
                    End If
                End If
            End If
        Next
    Next

    e3App.PutInfo 0, "Сортировка завершена. Начинаем переименование..."

    ' === Переименование по порядку ===
    For s = 1 To pooCountPlaced
        pooCounter = pooCounter + 1
        currentSymbolId = pooSymbolsToRename(s)(0)
        Dim newName : newName = "POO" & pooCounter

        symbol.SetId currentSymbolId
        symbol.SetName newName
    Next

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ: Переименовано " & pooCounter & " символов POO ==="
    Cleanup
End Sub

' === Очистка ===
Sub Cleanup
    ' Теперь эти переменные доступны, так как они объявлены глобально
    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' Вспомогательная процедура обмена элементов массива
Sub SwapArrayElements(arr, index1, index2)
    Dim tmp
    tmp = arr(index1)
    arr(index1) = arr(index2)
    arr(index2) = tmp
End Sub

' === Новая вспомогательная функция для извлечения номера листа ===
Function ExtractSheetNumber(sheetName)
    Dim re, matches
    Set re = New RegExp
    re.Pattern = "\d+" ' Ищем одну или более цифр
    re.Global = False ' Находим только первое совпадение
    
    Set matches = re.Execute(sheetName)
    
    If matches.Count > 0 Then
        ExtractSheetNumber = CInt(matches(0).Value) ' Преобразуем найденную строку в целое число
    Else
        ExtractSheetNumber = 0 ' Возвращаем 0, если число не найдено (или другая логика по умолчанию)
    End If
End Function

' === Запуск ===
Call RenamePOOSymbolsByLocation()