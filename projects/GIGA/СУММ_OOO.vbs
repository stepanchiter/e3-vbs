Option Explicit

' === Главная процедура === Суммирование атрибутов символов OOO
Sub SumOOOSymbolAttributes()
    Dim e3App, job, symbol
    Dim allSymbolIds(), allSymbolCount ' Переменные для всех символов в проекте
    Dim currentSymbolId, symbolName
    Dim s

    ' Переменные для хранения сумм атрибутов
    Dim totalEInom : totalEInom = 0.0
    Dim totalEIras : totalEIras = 0.0
    Dim totalEPnom : totalEPnom = 0.0
    Dim totalEPras : totalEPras = 0.0

    Dim attrValue ' Для временного хранения значения атрибута

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Суммирование атрибутов OOO символов ==="

    ' NEW: Получаем все ID символов в проекте напрямую
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "В проекте нет символов для анализа."
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    e3App.PutInfo 0, "Проверяем " & allSymbolCount & " символов на наличие OOO..."
    Dim oooFoundCount : oooFoundCount = 0

    ' Итерируем по всем символам, полученным из job.GetSymbolIds
    For s = 1 To allSymbolCount
        currentSymbolId = allSymbolIds(s)
        symbol.SetId(currentSymbolId)
        symbolName = symbol.GetName()

        ' Проверяем, является ли символ OOO
        If LCase(Left(symbolName, 3)) = "ooo" Then
            e3App.PutInfo 0, "  Найден OOO символ: " & symbolName & " (ID: " & currentSymbolId & ")"
            oooFoundCount = oooFoundCount + 1

            ' --- Чтение и суммирование атрибутов ---

            ' Атрибут ОД E_Inom
            attrValue = symbol.GetAttributeValue("ОД E_Inom")
            If IsNumeric(attrValue) And Len(Trim(attrValue)) > 0 Then
                totalEInom = totalEInom + CDbl(attrValue)
                e3App.PutInfo 0, "    Добавлено ОД E_Inom: " & attrValue
            Else
                e3App.PutInfo 0, "    ОД E_Inom: <пусто> или не число."
            End If

            ' Атрибут ОД E_Iras
            attrValue = symbol.GetAttributeValue("ОД E_Iras")
            If IsNumeric(attrValue) And Len(Trim(attrValue)) > 0 Then
                totalEIras = totalEIras + CDbl(attrValue)
                e3App.PutInfo 0, "    Добавлено ОД E_Iras: " & attrValue
            Else
                e3App.PutInfo 0, "    ОД E_Iras: <пусто> или не число."
            End If

            ' Атрибут ОД E_Pnom
            attrValue = symbol.GetAttributeValue("ОД E_Pnom")
            If IsNumeric(attrValue) And Len(Trim(attrValue)) > 0 Then
                totalEPnom = totalEPnom + CDbl(attrValue)
                e3App.PutInfo 0, "    Добавлено ОД E_Pnom: " & attrValue
            Else
                e3App.PutInfo 0, "    ОД E_Pnom: <пусто> или не число."
            End If

            ' Атрибут ОД E_Pras
            attrValue = symbol.GetAttributeValue("ОД E_Pras")
            If IsNumeric(attrValue) And Len(Trim(attrValue)) > 0 Then
                totalEPras = totalEPras + CDbl(attrValue)
                e3App.PutInfo 0, "    Добавлено ОД E_Pras: " & attrValue
            Else
                e3App.PutInfo 0, "    ОД E_Pras: <пусто> или не число."
            End If
        End If ' End If LCase(Left(symbolName, 3)) = "ooo"
    Next ' End For s = 1 To allSymbolCount

    e3App.PutInfo 0, "=== Суммарные значения атрибутов OOO символов ==="
    e3App.PutInfo 0, "Общее количество OOO символов: " & oooFoundCount
    e3App.PutInfo 0, "Сумма ОД E_Inom: " & totalEInom
    e3App.PutInfo 0, "Сумма ОД E_Iras: " & totalEIras
    e3App.PutInfo 0, "Сумма ОД E_Pnom: " & totalEPnom
    e3App.PutInfo 0, "Сумма ОД E_Pras: " & totalEPras
    e3App.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="

    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === Основной запуск ===
Call SumOOOSymbolAttributes()
