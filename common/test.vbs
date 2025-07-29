' Универсальный скрипт для соединения всех пинов заданной цепи - ВЕРСИЯ 2
' Обрабатывает все пины в проекте, включая поиск через connections

Set app = CreateObject("CT.Application") 
Set job = app.CreateJobObject()
Set symbol = job.CreateSymbolObject()
Set pin = job.CreatePinObject()
Set device = job.CreateDeviceObject()
Set connection = job.CreateConnectionObject()

' Словарь для хранения найденных пинов
Set dictPinIds = CreateObject("Scripting.Dictionary")

' Получение параметров от пользователя
wiregroupName = InputBox("Тип провода", "", "ПУГВнг(А)-LS")
If wiregroupName = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If

databaseWireName = InputBox("Сечение и цвет", "", "1х0.75(синий)")
If databaseWireName = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If

wireName = InputBox("Имя провода", "", "N")
If wireName = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If

signalName = InputBox("Имя цепи для поиска", "", "N")
If signalName = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If

app.PutInfo 0, "Начинаем поиск пинов цепи: " & signalName

' Основная функция поиска всех пинов цепи в проекте
Call FIND_ALL_PINS_BY_SIGNAL()

' Поиск пинов через connections
Call FIND_PINS_BY_CONNECTIONS()

' Создание соединений между найденными пинами
Call CREATE_CONNECTIONS()

app.PutInfo 0, "Обработка завершена."

'==========================================
' Функция поиска всех пинов по имени цепи
'==========================================
Sub FIND_ALL_PINS_BY_SIGNAL()
    app.PutInfo 0, "Поиск всех пинов в проекте по signal name..."
    
    ' Получаем все устройства в проекте
    deviceCount = job.GetDeviceIds(deviceIds)
    app.PutInfo 0, "Найдено устройств в проекте: " & deviceCount
    
    ' Используем Collection для динамического хранения данных пинов
    Set pinCollection = CreateObject("Scripting.Dictionary")
    Dim pinCounter
    pinCounter = 0
    
    ' Проходим по всем устройствам
    For deviceIndex = 1 To deviceCount
        deviceId = device.SetId(deviceIds(deviceIndex))
        deviceName = device.GetName()
        
        ' Получаем все пины устройства
        pinCount = device.GetPinIds(pinIds)
        
        If pinCount > 0 Then
            For pinIndex = 1 To pinCount
                pinId = pin.SetId(pinIds(pinIndex))
                pinName = pin.GetName()
                pinSignalName = pin.GetSignalName()
                
                ' Проверяем, соответствует ли пин нашей цепи
                If pinSignalName = signalName Then
                    ' Проверяем, не добавлен ли уже этот пин
                    If Not dictPinIds.Exists(pinId) Then
                        dictPinIds.Add pinId, deviceName & "." & pinName
                        app.PutInfo 0, "Найден пин (signal): " & deviceName & "." & pinName & " (ID: " & pinId & ") цепь: " & pinSignalName
                    End If
                End If
            Next
        End If
    Next
    
    app.PutInfo 0, "Найдено пинов по signal name: " & dictPinIds.Count
End Sub

'==========================================
' Функция поиска пинов через connections
'==========================================
Sub FIND_PINS_BY_CONNECTIONS()
    app.PutInfo 0, "Поиск пинов через connections..."
    
    ' Получаем все connections в проекте
    connectionCount = job.GetConnectionIds(connectionIds)
    app.PutInfo 0, "Найдено connections в проекте: " & connectionCount
    
    ' Множество для хранения ID пинов, связанных с нашей цепью
    Set connectedPinIds = CreateObject("Scripting.Dictionary")
    
    ' Сначала находим connections с нужной цепью
    For connectionIndex = 1 To connectionCount
        connectionId = connection.SetId(connectionIds(connectionIndex))
        connectionSignalName = connection.GetSignalName()
        
        If connectionSignalName = signalName Then
            ' Получаем пины этого connection
            pinCount = connection.GetPinIds(pinIds)
            
            For pinIndex = 1 To pinCount
                pinId = pinIds(pinIndex)
                connectedPinIds(pinId) = True
            Next
            
            app.PutInfo 0, "Connection ID: " & connectionId & " содержит " & pinCount & " пинов цепи " & signalName
        End If
    Next
    
    app.PutInfo 0, "Найдено пинов в connections: " & connectedPinIds.Count
    
    ' Теперь получаем информацию об этих пинах
    connectedPinKeys = connectedPinIds.Keys()
    For Each pinId In connectedPinKeys
        If Not dictPinIds.Exists(pinId) Then
            ' Получаем информацию о пине
            pin.SetId pinId
            pinName = pin.GetName()
            
            ' Получаем устройство, к которому принадлежит пин
            deviceId = device.SetId(pinId)
            deviceName = device.GetName()
            
            dictPinIds.Add pinId, deviceName & "." & pinName
            app.PutInfo 0, "Найден пин (connection): " & deviceName & "." & pinName & " (ID: " & pinId & ")"
        End If
    Next
    
    app.PutInfo 0, "Всего найдено уникальных пинов: " & dictPinIds.Count
End Sub

'==========================================
' Функция создания соединений
'==========================================
Sub CREATE_CONNECTIONS()
    If dictPinIds.Count < 2 Then
        app.PutInfo 0, "Недостаточно пинов для создания соединений (найдено: " & dictPinIds.Count & ")"
        Exit Sub
    End If
    
    app.PutInfo 0, "Создание соединений между пинами..."
    
    ' Сортируем пины перед созданием соединений
    Call SORT_FOUND_PINS()
    
    ' Получаем массив ключей (ID пинов)
    pinKeys = dictPinIds.Keys()
    
    ' Создаем соединения между соседними пинами
    For i = 0 To dictPinIds.Count - 2
        firstPinId = pinKeys(i)
        secondPinId = pinKeys(i + 1)
        
        firstPinName = dictPinIds(firstPinId)
        secondPinName = dictPinIds(secondPinId)
        
        app.PutInfo 0, "Соединяем: " & firstPinName & " -> " & secondPinName
        
        ' Создаем провод
        If CREATE_WIRE(firstPinId, secondPinId) Then
            app.PutInfo 0, "Соединение создано успешно"
        Else
            app.PutError 0, "Ошибка создания соединения"
        End If
    Next
    
    app.PutInfo 0, "Создано соединений: " & (dictPinIds.Count - 1)
End Sub

'==========================================
' Функция сортировки найденных пинов
'==========================================
Sub SORT_FOUND_PINS()
    If dictPinIds.Count <= 1 Then Exit Sub

    app.PutInfo 0, "Сортировка найденных пинов..."

    Dim pinKeys, pinCount, sortArray()
    pinKeys = dictPinIds.Keys()
    pinCount = dictPinIds.Count
    ReDim sortArray(pinCount - 1)

    Dim i
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

        ' Каждая строка — это массив из 3 элементов: [device, pin, pinId]
        sortArray(i) = Array(deviceName, pinName, pinId)
    Next

    ' Настройки сортировки
    ReDim options(1, 1)
    options(0, 0) = 0 ' индекс в sortArray(i)(0) — deviceName
    options(0, 1) = 2 ' инженерная сортировка
    options(1, 0) = 1 ' индекс sortArray(i)(1) — pinName
    options(1, 1) = 2 ' инженерная сортировка

    ' Сортировка
    app.SortArrayByIndexEx sortArray, options

    ' Очистка словаря и пересоздание
    dictPinIds.RemoveAll

    app.PutInfo 0, "Результат сортировки:"
    For i = 0 To pinCount - 1
        deviceName = sortArray(i)(0)
        pinName = sortArray(i)(1)
        pinId = sortArray(i)(2)
        dictPinIds.Add pinId, deviceName & "." & pinName
        app.PutInfo 0, (i + 1) & ". " & deviceName & "." & pinName & " (ID: " & pinId & ")"
    Next
End Sub


'==========================================
' Функция создания провода между двумя пинами
'==========================================
Function CREATE_WIRE(firstPinId, secondPinId)
    CREATE_WIRE = False
    
    If firstPinId <= 0 Or secondPinId <= 0 Then
        app.PutError 0, "Неверные ID пинов: " & firstPinId & ", " & secondPinId
        Exit Function
    End If
    
    ' Находим кабель (wire group) для создания провода
    cableCount = job.GetCableIds(cableIds)
    
    If cableCount = 0 Then
        app.PutError 0, "Не найдены кабели в проекте"
        Exit Function
    End If
    
    ' Ищем подходящий wire group
    For cableIndex = 1 To cableCount
        cableId = device.SetId(cableIds(cableIndex))
        cableName = device.GetName()
        isWireGroup = device.isWireGroup()
        
        If isWireGroup = 1 Then
            ' Создаем провод
            result = pin.CreateWire(wireName, wiregroupName, databaseWireName, cableId, 0, 0)
            
            If result > 0 Then
                ' Подключаем концы провода к пинам
                pin.SetEndPinId 1, firstPinId
                pin.SetEndPinId 2, secondPinId
                
                actualWireName = pin.GetName()
                app.PutInfo 0, "Создан провод: " & actualWireName & " (" & wiregroupName & ", " & databaseWireName & ")"
                
                CREATE_WIRE = True
                Exit Function
            Else
                app.PutError 0, "Ошибка создания провода в кабеле: " & cableName
            End If
        End If
    Next
    
    app.PutError 0, "Не найден подходящий wire group для создания провода"
End Function

' Освобождение ресурсов
Set dictPinIds = Nothing
Set connection = Nothing
Set device = Nothing
Set pin = Nothing
Set symbol = Nothing
Set job = Nothing 
Set app = Nothing