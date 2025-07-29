Set e3App = CreateObject("CT.Application")
Set job = e3App.CreateJobObject()

Set device = job.CreateDeviceObject()
Set symbol = job.CreateSymbolObject()
Set pin = job.CreatePinObject()

' Получаем ID всех выделенных символов
selectedCount = job.GetSelectedSymbolIds(selectedSymbolIds)
If selectedCount = 0 Then
    e3App.PutInfo 0, "Нет выделенных символов на схеме"
    WScript.Quit
End If

' Ограничиваем обработку одним символом (первым выделенным)
selectedSymbolId = selectedSymbolIds(1)
symbol.SetId(selectedSymbolId)
symbolRealId = symbol.GetId()

e3App.PutInfo 0, "Обработка символа ID: " & symbolRealId

' Ищем устройство, которому принадлежит символ
Dim foundDeviceId
foundDeviceId = 0

' Получаем все устройства в проекте
deviceCount = job.GetAllDeviceIds(deviceIds)
If deviceCount = 0 Then
    e3App.PutInfo 0, "В проекте нет устройств"
    WScript.Quit
End If

' Поиск устройства по символу
For i = 1 To deviceCount
    device.SetId(deviceIds(i))
    
    ' Получаем все символы устройства (3 - все символы, включая атрибутные)
    result = device.GetSymbolIds(symbolIds, 3)
    
    If result > 0 Then
        For j = 1 To result
            symbol.SetId(symbolIds(j))
            currentSymbolId = symbol.GetId()
            
            If currentSymbolId = symbolRealId Then
                foundDeviceId = deviceIds(i)
                Exit For
            End If
        Next
    End If
    
    If foundDeviceId <> 0 Then Exit For
Next

If foundDeviceId = 0 Then
    e3App.PutInfo 0, "Устройство для символа " & symbolRealId & " не найдено"
    WScript.Quit
End If

' Начинаем обработку устройства
device.SetId(foundDeviceId)
deviceName = device.GetName()
e3App.PutInfo 0, "Найдено устройство: " & deviceName & " (ID: " & foundDeviceId & ")"
e3App.PutInfo 0, "Очистка сигналов на пинах..."

' Получаем все пины устройства
pinCount = device.GetPinIds(pinIds)
If pinCount = 0 Then
    e3App.PutInfo 0, "У устройства нет пинов"
    WScript.Quit
End If

Dim clearedCount
clearedCount = 0

' Очищаем сигналы на всех пинах
For i = 1 To pinCount
    pin.SetId(pinIds(i))
    pinName = pin.GetName()
    
    ' Проверяем текущий сигнал
    currentSignal = pin.GetSignalName()
    If Len(currentSignal) > 0 Then
        ' Очищаем сигнал
        result = pin.SetSignalName("")
        If result = 1 Then
            clearedCount = clearedCount + 1
            e3App.PutInfo 0, "Очищен пин " & pinName & " (был сигнал: '" & currentSignal & "')"
        Else
            e3App.PutInfo 1, "Ошибка очистки пина " & pinName
        End If
    End If
Next

' Итоговый отчет
e3App.PutInfo 0, "Готово. Очищено сигналов: " & clearedCount & " из " & pinCount & " пинов"

' Очистка объектов
Set pin = Nothing
Set symbol = Nothing
Set device = Nothing
Set job = Nothing
Set e3App = Nothing