Set e3Application = CreateObject("CT.Application")
Set job = e3Application.CreateJobObject()

Set device = job.CreateDeviceObject()
Set symbol = job.CreateSymbolObject()

Dim deletedDevices
Set deletedDevices = CreateObject("Scripting.Dictionary")

' Получаем ID всех выделенных символов
selectedCount = job.GetSelectedSymbolIds(selectedSymbolIds)
If selectedCount = 0 Then
    e3Application.PutInfo 0, "Нет выделенных символов на схеме"
    WScript.Quit
End If

e3Application.PutInfo 0, "Выделено символов: " & selectedCount

' Получаем все устройства в проекте
deviceCount = job.GetAllDeviceIds(deviceIds)
If deviceCount = 0 Then
    e3Application.PutInfo 0, "В проекте нет устройств"
    WScript.Quit
End If

' Для каждого выделенного символа
For selectedIndex = 1 To selectedCount
    selectedSymbolId = selectedSymbolIds(selectedIndex)
    symbol.SetId(selectedSymbolId)
    selectedRealId = symbol.GetId()
    
    found = False

    ' Проходим по всем девайсам
    For i = 1 To deviceCount
        deviceId = deviceIds(i)

        If Not deletedDevices.Exists(CStr(deviceId)) Then
            device.SetId(deviceId)
            
            result = device.GetSymbolIds(symbolIds, 3) ' 3 — все символы, включая атрибутные
            
            If result > 0 Then
                For j = 1 To result
                    symbol.SetId(symbolIds(j))
                    currentSymbolId = symbol.GetId()
                    
                    If currentSymbolId = selectedRealId Then
                        found = True
                        deviceName = device.GetName()
                        
                        delResult = device.DeleteForced()
                        
                        If delResult = 1 Then
                            deletedDevices.Add CStr(deviceId), True
                            e3Application.PutInfo 0, "Удалено устройство: " & deviceName & " (" & deviceId & ")"
                        Else
                            e3Application.PutInfo 0, "Ошибка при удалении: " & deviceName & " (" & deviceId & ")"
                        End If
                        Exit For
                    End If
                Next
            End If

            If found Then Exit For
        End If
    Next

    If Not found Then
        e3Application.PutInfo 0, "Устройство для символа " & selectedRealId & " не найдено"
    End If
Next

' Очистка
Set deletedDevices = Nothing
Set symbol = Nothing
Set device = Nothing
Set job = Nothing
Set e3Application = Nothing
