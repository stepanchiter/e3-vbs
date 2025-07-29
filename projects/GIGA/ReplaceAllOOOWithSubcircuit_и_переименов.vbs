' === Главная процедура === Наложение фрагмента поверх всех символов OOO...
Sub ReplaceAllOOOWithSubcircuit()
    Dim e3App, job, symbol, sheet, device
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName, symbolIndex
    Dim s
    Dim subcircuitPath, subcircuitVersion
    Dim insertedDeviceIds(), deviceCount, d, devName, newName

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    Set sheet = job.CreateSheetObject()
    Set device = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Обработка всех символов OOO... ==="

    subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\SIC-tQF.e3p"
    subcircuitVersion = "1"

    e3App.PutInfo 0, "Используется фрагмент: " & subcircuitPath & " (версия: " & subcircuitVersion & ")"

    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount > 0 Then
        For s = 1 To allSymbolCount
            currentSymbolId = allSymbolIds(s)
            symbol.SetId(currentSymbolId)
            symbolName = symbol.GetName()

            If UCase(Left(symbolName, 3)) = "OOO" Then
                symbolIndex = Mid(symbolName, 4) ' Получаем индекс после "OOO"
                
                Dim oooX, oooY, oooSheetId, gridDesc, colVal, rowVal
                oooSheetId = symbol.GetSchemaLocation(oooX, oooY, gridDesc, colVal, rowVal)

                If oooSheetId > 0 Then
                    sheet.SetId oooSheetId
                    e3App.PutInfo 0, "Найден символ " & symbolName & " на листе ID: " & oooSheetId & ", координаты (" & oooX & ", " & oooY & ")"

                    ' Вставка фрагмента
                    Dim result : result = sheet.PlacePart(subcircuitPath, subcircuitVersion, oooX, oooY, 0.0)

                    If result = 0 Or result = -3 Then
                        e3App.PutInfo 0, "Фрагмент установлен для " & symbolName

                        ' Поиск всех устройств, созданных после вставки
                        deviceCount = job.GetDeviceIds(insertedDeviceIds)
                        For d = 1 To deviceCount
                            device.SetId insertedDeviceIds(d)
                            devName = device.GetName()
                            If LCase(Left(devName, 4)) = "-tqf" Then
                                newName = "-QF" & symbolIndex
                                e3App.PutInfo 0, "Переименование девайса " & devName & " > " & newName
                                device.SetName newName
                                Exit For ' предполагаем, что только один девайс создаётся
                            End If
                        Next
                    Else
                        e3App.PutInfo 0, "Ошибка установки фрагмента для " & symbolName & ". Код: " & result
                    End If
                Else
                    e3App.PutInfo 0, "Символ " & symbolName & " не размещён на схеме."
                End If
            End If
        Next
    Else
        e3App.PutInfo 0, "Символы в проекте не найдены."
    End If

    ' Очистка
    Set device = Nothing
    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

Call ReplaceAllOOOWithSubcircuit()
