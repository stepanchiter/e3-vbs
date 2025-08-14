'===============================================================================
' СКРИПТ АВТОМАТИЧЕСКОГО СОЕДИНЕНИЯ ПИНОВ ПО ЦЕПИ В E3.SERIES
'===============================================================================
' Описание: Автоматически находит все пины в проекте e3.series, принадлежащие
'           к указанной цепи (signal name), сортирует их и создает между ними
'           последовательные соединения проводами.
'
' Входные параметры:
'   - wiregroupName     : Имя группы проводов
'   - databaseWireName  : Имя провода в базе данных
'   - wireName          : Имя создаваемого провода
'   - signalName        : Имя цепи для поиска пинов
'
' Алгоритм работы:
'   1. Поиск всех пинов в проекте по имени сигнала (SignalName)
'   2. Дополнительный поиск пинов через объекты Connection
'   3. Сортировка найденных пинов по имени устройства и пина
'   4. Создание последовательных соединений между соседними пинами
'
' Автор: [Ваше имя]
' Дата:  [Дата создания]
'===============================================================================
Option Explicit

' --- Получение параметров от HTA ---
If WScript.Arguments.Count <> 4 Then
    MsgBox "Ожидается 4 параметра: wiregroupName, databaseWireName, wireName, signalName", vbCritical, "Ошибка запуска"
    WScript.Quit 1
End If

Dim wiregroupName, databaseWireName, wireName, signalName
wiregroupName     = WScript.Arguments(0)
databaseWireName  = WScript.Arguments(1)
wireName          = WScript.Arguments(2)
signalName        = WScript.Arguments(3)

' --- Инициализация E3 API ---
Dim app, job, symbol, pin, device, connection
Set app        = CreateObject("CT.Application")
Set job        = app.CreateJobObject()
Set symbol     = job.CreateSymbolObject()
Set pin        = job.CreatePinObject()
Set device     = job.CreateDeviceObject()
Set connection = job.CreateConnectionObject()

Dim dictPinIds
Set dictPinIds = CreateObject("Scripting.Dictionary")

app.PutInfo 0, "Начинаем поиск пинов цепи: " & signalName

Call FIND_ALL_PINS_BY_SIGNAL()
Call FIND_PINS_BY_CONNECTIONS()
Call CREATE_CONNECTIONS()

app.PutInfo 0, "Обработка завершена."

' === Поиск пинов по SignalName ===
Sub FIND_ALL_PINS_BY_SIGNAL()
    app.PutInfo 0, "Поиск всех пинов в проекте по signal name..."

    Dim deviceIds, pinIds
    Dim deviceCount, pinCount, deviceIndex, pinIndex, pinId, deviceId
    deviceCount = job.GetDeviceIds(deviceIds)
    app.PutInfo 0, "Найдено устройств в проекте: " & deviceCount

    For deviceIndex = 1 To deviceCount
        deviceId = device.SetId(deviceIds(deviceIndex))
        Dim deviceName
        deviceName = device.GetName()

        pinCount = device.GetPinIds(pinIds)
        If pinCount > 0 Then
            For pinIndex = 1 To pinCount
                pinId = pin.SetId(pinIds(pinIndex))
                Dim pinName, pinSignalName
                pinName = pin.GetName()
                pinSignalName = pin.GetSignalName()

                If pinSignalName = signalName Then
                    If Not dictPinIds.Exists(pinId) Then
                        dictPinIds.Add pinId, deviceName & "." & pinName
                        app.PutInfo 0, "Найден пин (signal): " & deviceName & "." & pinName & " (ID: " & pinId & ")"
                    End If
                End If
            Next
        End If
    Next

    app.PutInfo 0, "Найдено пинов по signal name: " & dictPinIds.Count
End Sub

' === Поиск через connections ===
Sub FIND_PINS_BY_CONNECTIONS()
    app.PutInfo 0, "Поиск пинов через connections..."

    Dim connectionIds, pinIds, connectionIndex, connectionId
    Dim connectedPinIds
    Set connectedPinIds = CreateObject("Scripting.Dictionary")

    Dim connectionCount
    connectionCount = job.GetConnectionIds(connectionIds)
    app.PutInfo 0, "Найдено connections в проекте: " & connectionCount

    For connectionIndex = 1 To connectionCount
        connectionId = connection.SetId(connectionIds(connectionIndex))
        If connection.GetSignalName() = signalName Then
            Dim pinCount, pinIndex, pinId
            pinCount = connection.GetPinIds(pinIds)
            For pinIndex = 1 To pinCount
                pinId = pinIds(pinIndex)
                connectedPinIds(pinId) = True
            Next
            app.PutInfo 0, "Connection ID: " & connectionId & " содержит " & pinCount & " пинов цепи " & signalName
        End If
    Next

    Dim connectedPinKeys, key
    connectedPinKeys = connectedPinIds.Keys()
    For Each key In connectedPinKeys
        If Not dictPinIds.Exists(key) Then
            pin.SetId key
            Dim pinName, devName
            pinName = pin.GetName()
            device.SetId key
            devName = device.GetName()
            dictPinIds.Add key, devName & "." & pinName
            app.PutInfo 0, "Найден пин (connection): " & devName & "." & pinName & " (ID: " & key & ")"
        End If
    Next

    app.PutInfo 0, "Всего найдено уникальных пинов: " & dictPinIds.Count
End Sub

' === Создание соединений ===
Sub CREATE_CONNECTIONS()
    If dictPinIds.Count < 2 Then
        app.PutInfo 0, "Недостаточно пинов для создания соединений (найдено: " & dictPinIds.Count & ")"
        Exit Sub
    End If

    app.PutInfo 0, "Создание соединений между пинами..."
    Call SORT_FOUND_PINS()

    Dim pinKeys, i, firstPinId, secondPinId
    pinKeys = dictPinIds.Keys()

    For i = 0 To dictPinIds.Count - 2
        firstPinId  = pinKeys(i)
        secondPinId = pinKeys(i + 1)

        app.PutInfo 0, "Соединяем: " & dictPinIds(firstPinId) & " -> " & dictPinIds(secondPinId)
        If CREATE_WIRE(firstPinId, secondPinId) Then
            app.PutInfo 0, "Соединение создано успешно"
        Else
            app.PutError 0, "Ошибка создания соединения"
        End If
    Next

    app.PutInfo 0, "Создано соединений: " & (dictPinIds.Count - 1)
End Sub

' === Сортировка найденных пинов ===
Sub SORT_FOUND_PINS()
    If dictPinIds.Count <= 1 Then Exit Sub

    app.PutInfo 0, "Сортировка найденных пинов..."

    Dim pinKeys, pinCount, sortArray(), i
    pinKeys = dictPinIds.Keys()
    pinCount = dictPinIds.Count
    ReDim sortArray(pinCount - 1)

    For i = 0 To pinCount - 1
        Dim fullName, deviceName, pinName, dotPos, pinId
        pinId = pinKeys(i)
        fullName = CStr(dictPinIds(pinId))
        dotPos = InStr(fullName, ".")

        If dotPos > 0 Then
            deviceName = Left(fullName, dotPos - 1)
            pinName = Mid(fullName, dotPos + 1)
        Else
            deviceName = fullName
            pinName = ""
        End If

        sortArray(i) = Array(deviceName, pinName, pinId)
    Next

    Dim options(1, 1)
    options(0, 0) = 0 ' deviceName
    options(0, 1) = 2
    options(1, 0) = 1 ' pinName
    options(1, 1) = 2

    app.SortArrayByIndexEx sortArray, options

    dictPinIds.RemoveAll

    For i = 0 To pinCount - 1
        deviceName = sortArray(i)(0)
        pinName = sortArray(i)(1)
        pinId = sortArray(i)(2)
        dictPinIds.Add pinId, deviceName & "." & pinName
        app.PutInfo 0, (i + 1) & ". " & deviceName & "." & pinName & " (ID: " & pinId & ")"
    Next
End Sub

' === Создание провода между двумя пинами ===
Function CREATE_WIRE(firstPinId, secondPinId)
    CREATE_WIRE = False

    If firstPinId <= 0 Or secondPinId <= 0 Then
        app.PutError 0, "Неверные ID пинов: " & firstPinId & ", " & secondPinId
        Exit Function
    End If

    Dim cableCount, cableIds, cableId, cableName
    cableCount = job.GetCableIds(cableIds)

    If cableCount = 0 Then
        app.PutError 0, "Не найдены кабели в проекте"
        Exit Function
    End If

    Dim i, result, actualWireName
    For i = 1 To cableCount
        cableId = device.SetId(cableIds(i))
        If device.IsWireGroup() = 1 Then
            result = pin.CreateWire(wireName, wiregroupName, databaseWireName, cableId, 0, 0)
            If result > 0 Then
                pin.SetEndPinId 1, firstPinId
                pin.SetEndPinId 2, secondPinId
                actualWireName = pin.GetName()
                app.PutInfo 0, "Создан провод: " & actualWireName & " (" & wiregroupName & ", " & databaseWireName & ")"
                CREATE_WIRE = True
                Exit Function
            Else
                cableName = device.GetName()
                app.PutError 0, "Ошибка создания провода в кабеле: " & cableName
            End If
        End If
    Next

    app.PutError 0, "Не найден подходящий wire group для создания провода"
End Function

' === Очистка ===
Set dictPinIds = Nothing
Set connection = Nothing
Set device     = Nothing
Set pin        = Nothing
Set symbol     = Nothing
Set job        = Nothing
Set app        = Nothing
