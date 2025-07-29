Option Explicit

' =========================================================================
' Вспомогательная функция: Извлечение числового индекса из имени
' Пример: "POO19" -> 19, "-XT19" -> 19, "-tXT1" -> 1, "-QF12" -> 12
' Возвращает 0, если числовой индекс не найден или имя пустое/некорректное
' =========================================================================
Function ExtractNumericIndex(nameString)
    ' Включаем локальную обработку ошибок для этой функции
    On Error Resume Next
    
    Dim i, char, numericPart
    numericPart = ""
    ExtractNumericIndex = 0 ' Устанавливаем значение по умолчанию

    If Len(nameString) = 0 Then
        Exit Function ' Выходим, если строка пустая
    End If

    ' Начинаем поиск цифр с конца строки, так как они всегда в конце имени
    For i = Len(nameString) To 1 Step -1
        char = Mid(nameString, i, 1)
        If IsNumeric(char) Then
            numericPart = char & numericPart ' Добавляем цифру в начало numericPart
        Else
            ' Если встретили нецифровой символ, и уже есть цифры, значит, число закончилось
            If Len(numericPart) > 0 Then Exit For
        End If
    Next

    If Len(numericPart) > 0 Then
        If IsNumeric(numericPart) Then
            ExtractNumericIndex = CInt(numericPart)
        Else
            e3App.PutInfo 2, "Внутренняя ошибка в ExtractNumericIndex: Получена нечисловая часть '" & numericPart & "' из '" & nameString & "'"
        End If
    End If
    
    If Err.Number <> 0 Then
        e3App.PutInfo 2, "Ошибка в ExtractNumericIndex для строки '" & nameString & "': " & Err.Description
        Err.Clear ' Очищаем ошибку
    End If
End Function


' =========================================================================
' Основная процедура: Копирование атрибута E_TAG из POOi в Функцию -QFi и -KMi
' =========================================================================
Sub CopyETagToFunctionAttribute()
    On Error Resume Next ' Включаем обработку ошибок для всей процедуры

    ' Инициализация объектов E3.series
    Dim e3App, job, device, symbol
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set device = job.CreateDeviceObject()
    Set symbol = job.CreateSymbolObject()

    ' Словарь для хранения соответствия: Индекс POO -> Значение E_TAG
    Dim pooIndexToETagMap
    Set pooIndexToETagMap = CreateObject("Scripting.Dictionary")

    Dim i ' Переменная для циклов
    Dim allSymbolIds, symbolCount, currentSymbolId, currentSymbolName
    Dim attrValue, pooNumericIndex
    Dim processedDevicesCount ' Счетчик успешно обработанных устройств
    processedDevicesCount = 0

    e3App.PutInfo 0, "=== СТАРТ: Копирование атрибута E_TAG из POOi в Функцию -QFi и -KMi ==="

    ' =============================================================
    ' ФАЗА 1: Сбор значений атрибута "ОД E_TAG" из символов POOi
    ' =============================================================
    e3App.PutMessageEx 0, "Фаза 1: Сбор данных E_TAG из символов POOi...", 0, 0, 0, 249

    symbolCount = job.GetSymbolIds(allSymbolIds)

    If symbolCount > 0 Then
        For i = 1 To symbolCount
            currentSymbolId = allSymbolIds(i)
            symbol.SetId(currentSymbolId)
            currentSymbolName = symbol.GetName()

            If Left(UCase(currentSymbolName), 3) = "POO" Then ' Проверка на "POO"
                pooNumericIndex = ExtractNumericIndex(currentSymbolName)
                
                If pooNumericIndex > 0 Then ' Только если индекс извлечен корректно
                    attrValue = symbol.GetAttributeValue("ОД E_TAG")
                    
                    If attrValue <> "" Then
                        If Not pooIndexToETagMap.Exists(pooNumericIndex) Then
                            pooIndexToETagMap.Add pooNumericIndex, attrValue
                            e3App.PutInfo 0, "Найден символ POO" & pooNumericIndex & ". E_TAG: '" & attrValue & "'"
                        Else
                            ' Это предупреждение, если несколько символов POOi имеют один и тот же индекс,
                            ' но разные E_TAG. Берем первое найденное значение.
                            e3App.PutInfo 1, "Внимание: Дублирующийся индекс POO" & pooNumericIndex & ". Используется первое найденное значение E_TAG: '" & pooIndexToETagMap(pooNumericIndex) & "'. Текущее E_TAG: '" & attrValue & "'"
                        End If
                    Else
                        e3App.PutInfo 1, "ПРЕДУПРЕЖДЕНИЕ: Символ POO" & pooNumericIndex & " имеет пустой атрибут 'ОД E_TAG'. Пропускаю."
                    End If
                Else
                    e3App.PutInfo 1, "ПРЕДУПРЕЖДЕНИЕ: Не удалось извлечь числовой индекс из имени символа POO: '" & currentSymbolName & "'. Символ пропускается в Фазе 1."
                End If
            End If
        Next
    Else
        e3App.PutInfo 1, "В проекте не найдено ни одного символа POO."
    End If

    ' =============================================================
    ' ФАЗА 2: Поиск устройств -QFi и -KMi и копирование атрибута "Функция"
    ' Все сообщения (включая заголовок и итоговые) удалены из этой фазы.
    ' =============================================================
    Dim allDeviceIds, deviceCount, currentDeviceId, currentDeviceName
    Dim qfNumericIndex, eTagValueToSet, setResult
    Dim deviceNamePrefix ' Переменная для хранения префикса имени устройства

    deviceCount = job.GetAllDeviceIds(allDeviceIds) ' Получаем все устройства в проекте

    If deviceCount > 0 Then
        For i = 1 To deviceCount
            currentDeviceId = allDeviceIds(i)
            device.SetId(currentDeviceId)
            currentDeviceName = device.GetName()

            deviceNamePrefix = Left(UCase(currentDeviceName), 3)

            ' Проверяем, что устройство имеет ожидаемый формат имени "-QFi" или "-KMi"
            If deviceNamePrefix = "-QF" Or deviceNamePrefix = "-KM" Then
                qfNumericIndex = ExtractNumericIndex(currentDeviceName)

                If qfNumericIndex > 0 Then ' Только если индекс извлечен корректно
                    If pooIndexToETagMap.Exists(qfNumericIndex) Then
                        eTagValueToSet = pooIndexToETagMap(qfNumericIndex)
                        
                        ' Устанавливаем значение атрибута "Функция"
                        ' Если атрибута нет, E3.series его создаст
                        setResult = device.SetAttributeValue("Функция", eTagValueToSet)
                        
                        If setResult = 0 Then
                            processedDevicesCount = processedDevicesCount + 1
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' Сообщения удалены
    End If

    e3App.PutInfo 0, "=== ЗАВЕРШЕНО ==="

    ' Очистка объектов
    Set pooIndexToETagMap = Nothing
    Set symbol = Nothing
    Set device = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' Вызов основной процедуры
Call CopyETagToFunctionAttribute()