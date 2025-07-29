' Скрипт для поиска проблемных реле с возможностью перехода
Set e3App = CreateObject("CT.Application")
Set job = e3App.CreateJobObject()
Set device = job.CreateDeviceObject()
Set symbol = job.CreateSymbolObject()

' Цвета сообщений (RGB)
Const COLOR_RED = 225    ' Для ошибок
Const COLOR_GREEN = 121 ' Для успешных сообщений
Const COLOR_BLUE = 249 ' Для информационных

Dim deviceIds, deviceCount, devName, hasProblems
hasProblems = False

' Очищаем окно сообщений
e3App.PutMessageEx 0, "Поиск проблемных реле...", 0, 0, 0, COLOR_BLUE

' Получаем все устройства проекта
deviceCount = job.GetAllDeviceIds(deviceIds)

If deviceCount > 0 Then
    For i = 1 To deviceCount
        device.SetId(deviceIds(i))
        devName = device.GetName()
        
        ' Фильтруем реле (по "KL" в имени)
        If InStr(1, devName, "KL", vbTextCompare) > 0 Then
            Dim symbolIds, symbolCount, coilConnected, anyContactConnected
            coilConnected = False
            anyContactConnected = False
            
            ' Получаем все символы устройства
            symbolCount = device.GetSymbolIds(symbolIds, 0)
            
            If symbolCount > 0 Then
                ' Анализируем все символы реле
                For j = 1 To symbolCount
                    symbol.SetId(symbolIds(j))
                    symbolTypeName = LCase(symbol.GetSymbolTypeName())
                    
                    ' Проверяем тип символа и его подключение
                    If InStr(symbolTypeName, "катушка") > 0 Then
                        If symbol.IsConnected() = 1 Then
                            coilConnected = True
                        End If
                    ElseIf InStr(symbolTypeName, "контакт") > 0 Then
                        If symbol.IsConnected() = 1 Then
                            anyContactConnected = True
                        End If
                    End If
                Next
                
                ' Формируем сообщение о проблемах с кликабельной ссылкой
                If Not coilConnected Then
                    e3App.PutMessageEx 0, "РЕЛЕ " & devName & ": Катушка не подключена", deviceIds(i), COLOR_RED, 0, 0
                    hasProblems = True
                ElseIf Not anyContactConnected Then
                    e3App.PutMessageEx 0, "РЕЛЕ " & devName & ": Катушка подключена, но ни один контакт не подключен", deviceIds(i), COLOR_RED, 0, 0
                    hasProblems = True
                End If
            End If
        End If
    Next
    
    ' Выводим итоговое сообщение
    If Not hasProblems Then
        e3App.PutMessageEx 0, "Все реле подключены корректно.", 0, 0, COLOR_GREEN, 0
    End If
Else
    e3App.PutMessageEx 0, "В проекте нет устройств.", 0, COLOR_RED, 0, 0
End If

' Освобождаем ресурсы
Set symbol = Nothing
Set device = Nothing
Set job = Nothing
Set e3App = Nothing