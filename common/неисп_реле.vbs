' Скрипт для проверки подключения пинов у реле (устройства с "KL" в имени)
' Выводит только реле с полностью неподключенными катушками или контактами

Set e3App = CreateObject("CT.Application")
Set job = e3App.CreateJobObject()

Set device = job.CreateDeviceObject()
Set pin = job.CreatePinObject()

Dim relayCoilPins, relayContactPins
relayCoilPins = Array("A1", "A2")
relayContactPins = Array("11", "12", "14", "21", "22", "24", "31", "32", "34", "41", "42", "44")

Dim deviceIds, deviceCount

' Получаем все девайсы проекта
deviceCount = job.GetAllDeviceIds(deviceIds)

If deviceCount > 0 Then
    For i = 1 To deviceCount
        device.SetId(deviceIds(i))
        devName = device.GetName()

        ' Фильтрация реле по имени
        If InStr(devName, "KL") > 0 Then
            Dim pinIds, pinCount
            pinCount = device.GetPinIds(pinIds)
            
            Dim pinMap, pinSignalNames
            Set pinMap = CreateObject("Scripting.Dictionary")
            Set pinSignalNames = CreateObject("Scripting.Dictionary")
            
            ' Инициализация пинов катушки
            For Each p In relayCoilPins
                pinMap.Add p, False
                pinSignalNames.Add p, ""
            Next
            
            ' Инициализация только существующих контактных пинов
            If pinCount > 0 Then
                For j = 1 To pinCount
                    pin.SetId(pinIds(j))
                    pinName = pin.GetName()
                    
                    ' Добавляем только те контактные пины, которые есть у устройства
                    For Each contactPin In relayContactPins
                        If pinName = contactPin And Not pinMap.Exists(pinName) Then
                            pinMap.Add pinName, False
                            pinSignalNames.Add pinName, ""
                        End If
                    Next
                Next
            End If
            
            ' Проверка подключения пинов через GetSignalName()
            If pinCount > 0 Then
                For j = 1 To pinCount
                    pin.SetId(pinIds(j))
                    pinName = pin.GetName()

                    If pinMap.Exists(pinName) Then
                        signalName = pin.GetSignalName()
                        pinMap(pinName) = (Len("" & signalName) > 0)
                        pinSignalNames(pinName) = signalName
                    End If
                Next
            End If

            ' Проверка, что не подключены ОБА пина катушки
            Dim bothCoilsDisconnected
            bothCoilsDisconnected = True
            For Each p In relayCoilPins
                If pinMap.Exists(p) And pinMap(p) Then
                    bothCoilsDisconnected = False
                    Exit For
                End If
            Next
            
            ' Проверка, что не подключены ВСЕ контактные пины
            Dim allContactsDisconnected
            allContactsDisconnected = True
            For Each p In pinMap.Keys()
                ' Пропускаем пины катушки
                If Not IsInArray(p, relayCoilPins) Then
                    If pinMap(p) Then
                        allContactsDisconnected = False
                        Exit For
                    End If
                End If
            Next

            ' Вывод сообщений только для проблемных реле
            If bothCoilsDisconnected Or allContactsDisconnected Then
                e3App.PutInfo 0, "Реле " & devName & " - проблемы с подключением:"
                
                If bothCoilsDisconnected Then
                    e3App.PutInfo 0, "  • Не подключены ОБА пина катушки (A1 и A2)"
                End If
                
                If allContactsDisconnected Then
                    e3App.PutInfo 0, "  • Не подключены ВСЕ контактные пины"
                End If
                
                ' Дополнительная информация о неподключенных пинах
                Dim disconnectedPins
                disconnectedPins = ""
                For Each p In pinMap.Keys()
                    If Not pinMap(p) Then
                        disconnectedPins = disconnectedPins & p & ", "
                    End If
                Next
                If Len(disconnectedPins) > 0 Then
                    e3App.PutInfo 0, "  • Неподключенные пины: " & Left(disconnectedPins, Len(disconnectedPins)-2)
                End If
            End If
        End If
    Next
Else
    e3App.PutInfo 0, "Нет устройств в проекте"
End If

Set pin = Nothing
Set device = Nothing
Set job = Nothing
Set e3App = Nothing

' Вспомогательная функция для проверки наличия элемента в массиве
Function IsInArray(item, arr)
    IsInArray = False
    For Each element In arr
        If element = item Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function