Option Explicit

' === Главная процедура === Наложение фрагмента поверх символа OOO1
Sub ReplaceOOO1WithSubcircuit()
    Dim e3App, job, symbol, sheet
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName
    Dim s
    Dim ooo1Found : ooo1Found = False
    Dim ooo1SymbolId, ooo1SheetId
    Dim ooo1XPos, ooo1YPos
    Dim subcircuitPath, subcircuitVersion
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    Set sheet = job.CreateSheetObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Наложение фрагмента поверх символа OOO1 ==="

    ' Задаем путь к фрагменту и версию прямо в коде (можно изменить по необходимости)
    subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\SIC-QF666.e3p"  ' Измените на нужный путь или имя фрагмента
    subcircuitVersion = "1"             ' Измените на нужную версию
    
    e3App.PutInfo 0, "Используется фрагмент: " & subcircuitPath & " (версия: " & subcircuitVersion & ")"

    ' Получаем все ID символов в проекте
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount > 0 Then
        e3App.PutInfo 0, "Найдено " & allSymbolCount & " символов в проекте. Поиск символа OOO1..."

        ' === Поиск символа OOO1 ===
        For s = 1 To allSymbolCount
            currentSymbolId = allSymbolIds(s)
            symbol.SetId(currentSymbolId)
            symbolName = symbol.GetName()

            If UCase(symbolName) = "OOO1" Then
                Dim gridDesc, columnValue, rowValue
                ooo1SheetId = symbol.GetSchemaLocation(ooo1XPos, ooo1YPos, gridDesc, columnValue, rowValue)
                
                If ooo1SheetId > 0 Then ' Символ размещен на схеме
                    ooo1SymbolId = currentSymbolId
                    ooo1Found = True
                    
                    sheet.SetId ooo1SheetId
                    Dim sheetName : sheetName = sheet.GetName()
                    
                    e3App.PutInfo 0, "Символ OOO1 найден:"
                    e3App.PutInfo 0, "  ID: " & ooo1SymbolId
                    e3App.PutInfo 0, "  Лист: " & sheetName & " (ID: " & ooo1SheetId & ")"
                    e3App.PutInfo 0, "  Позиция: " & columnValue & rowValue & " (X: " & ooo1XPos & ", Y: " & ooo1YPos & ")"
                    Exit For
                Else
                    e3App.PutInfo 0, "Символ OOO1 найден, но НЕ размещен на схеме (ID: " & currentSymbolId & ")"
                End If
            End If
        Next

        If ooo1Found Then
            ' === Установка фрагмента поверх OOO1 ===
            e3App.PutInfo 0, "Установка фрагмента '" & subcircuitPath & "' поверх символа OOO1 в позицию (" & ooo1XPos & ", " & ooo1YPos & ")..."
            
            sheet.SetId ooo1SheetId
            Dim placeResult : placeResult = sheet.PlacePart(subcircuitPath, subcircuitVersion, ooo1XPos, ooo1YPos, 0.0)
            
            ' Обработка результата установки фрагмента
            Select Case placeResult
                Case 0
                    e3App.PutInfo 0, "Фрагмент '" & subcircuitPath & "' успешно установлен поверх символа OOO1 на лист " & sheet.GetName() & " в позицию (" & ooo1XPos & ", " & ooo1YPos & ")"
                Case 9
                    e3App.PutInfo 0, "ОШИБКА: Несовместимая версия файла фрагмента"
                Case 3
                    e3App.PutInfo 0, "ОШИБКА: Неверное имя фрагмента или версия"
                Case -1
                    e3App.PutInfo 0, "ОШИБКА: Фрагмент состоит из нескольких листов и установлен параметр 'Ignore sheet border'"
                Case -2
                    e3App.PutInfo 0, "ОШИБКА: Фрагмент содержит листы и НЕ установлен параметр 'Ignore sheet border'"
                Case -3
                    e3App.PutInfo 0, "ОШИБКА: Фрагмент уже размещен или конфликт объектов в позиции (" & ooo1XPos & ", " & ooo1YPos & ")"
                    e3App.PutInfo 0, "ПРИМЕЧАНИЕ: Это может быть нормально, если фрагмент наложился поверх символа OOO1"
                Case -4
                    e3App.PutInfo 0, "ОШИБКА: Лист заблокирован"
                Case Else
                    e3App.PutInfo 0, "ОШИБКА: Неизвестная ошибка при установке фрагмента. Код ошибки: " & placeResult
            End Select

            If placeResult = 0 Then
                e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА: Фрагмент успешно наложен поверх символа OOO1 ==="
            ElseIf placeResult = -3 Then
                e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА: Возможен конфликт размещения (код -3), но фрагмент может быть установлен ==="
            Else
                e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА: Ошибка при наложении фрагмента поверх символа OOO1 ==="
            End If
        Else
            e3App.PutInfo 0, "Символ OOO1 не найден в проекте или не размещен на схеме."
            e3App.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="
        End If
    Else
        e3App.PutInfo 0, "В проекте нет символов для анализа. Скрипт завершен."
        e3App.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="
    End If

    ' Очистка объектов
    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === Основной запуск ===
Call ReplaceOOO1WithSubcircuit()