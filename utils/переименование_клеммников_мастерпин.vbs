' Скрипт: Переименование устройств -sXT1 в -XT666 и их мастерпинов в 6 / 66
' Обрабатывает только устройства с существующими мастерпинами

Set app = CreateObject("CT.Application")
Set job = app.CreateJobObject()
Set device = job.CreateDeviceObject()

Dim deviceIds
result = job.GetAllDeviceIds(deviceIds)

Dim foundCount
foundCount = 0

app.PutInfo 0, "=== ПОИСК УСТРОЙСТВ -sXT1 ==="

If result > 0 Then
    app.PutInfo 0, "Всего устройств в проекте: " & result
    
    ' Сначала найдем все устройства -sXT1 и проверим их мастерпины
    Dim sxtDevices()
    Dim sxtCount
    sxtCount = 0
    
    For i = 1 To result
        device.SetId deviceIds(i)
        name = device.GetName()
        
        If name = "-sXT1" Then
            ' Проверяем, есть ли у устройства мастерпин
            currentMasterPin = device.GetMasterPinName()
            
            app.PutInfo 0, "--- Найдено устройство -sXT1 ---"
            app.PutInfo 0, "ID устройства: " & deviceIds(i)
            app.PutInfo 0, "Текущий мастерпин: '" & currentMasterPin & "'"
            
            ' Если мастерпин существует (не пустой)
            If Len(Trim(currentMasterPin)) > 0 Then
                sxtCount = sxtCount + 1
                ReDim Preserve sxtDevices(sxtCount - 1)
                sxtDevices(sxtCount - 1) = deviceIds(i)
                app.PutInfo 0, "Устройство добавлено в список для обработки (№" & sxtCount & ")"
            Else
                app.PutInfo 0, "Устройство пропущено - нет мастерпина"
            End If
        End If
    Next
    
    app.PutInfo 0, "=== ОБРАБОТКА НАЙДЕННЫХ УСТРОЙСТВ ==="
    app.PutInfo 0, "Устройств с мастерпинами для обработки: " & sxtCount
    
    ' Теперь обрабатываем найденные устройства с мастерпинами
    For j = 0 To sxtCount - 1
        If j >= 2 Then
            app.PutInfo 0, "Найдено более двух устройств с мастерпинами. Остальные не обработаны."
            Exit For
        End If
        
        device.SetId sxtDevices(j)
        foundCount = j + 1
        
        app.PutInfo 0, "--- Обработка устройства #" & foundCount & " ---"
        
        ' Переименование устройства
        newName = "-XT666"
        resultSet = device.SetName(newName)
        
        If resultSet = 0 Then
            app.PutInfo 0, "Ошибка при переименовании устройства #" & foundCount
        Else
            app.PutInfo 0, "Устройство #" & foundCount & " переименовано в " & newName
        End If
        
        ' Установка мастерпина
        If foundCount = 1 Then
            resultPin = device.SetMasterPinName("6")
            If resultPin = 0 Then
                app.PutInfo 0, "Ошибка при установке мастерпина '6' для устройства #" & foundCount
            Else
                app.PutInfo 0, "Мастерпин устройства #" & foundCount & " установлен в: 6"
            End If
        ElseIf foundCount = 2 Then
            resultPin = device.SetMasterPinName("66")
            If resultPin = 0 Then
                app.PutInfo 0, "Ошибка при установке мастерпина '66' для устройства #" & foundCount
            Else
                app.PutInfo 0, "Мастерпин устройства #" & foundCount & " установлен в: 66"
            End If
        End If
        
        ' Проверка результата
        finalMasterPin = device.GetMasterPinName()
        app.PutInfo 0, "Итоговый мастерпин: '" & finalMasterPin & "'"
    Next
    
    If sxtCount = 0 Then
        app.PutInfo 0, "Не найдено ни одного устройства -sXT1 с мастерпином."
    End If
    
Else
    app.PutInfo 0, "Ошибка: устройства в проекте не найдены."
End If

app.PutInfo 0, "=== СКРИПТ ЗАВЕРШЕН ==="

' Очистка
Set device = Nothing
Set job = Nothing
Set app = Nothing