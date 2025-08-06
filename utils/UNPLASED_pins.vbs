Option Explicit

' === Константы ===
Const COLOR_RED = 225
Const COLOR_GREEN = 121
Const COLOR_BLUE = 249

' === Инициализация приложения ===
Dim e3App : Set e3App = CreateObject("CT.Application")
Dim job : Set job = e3App.CreateJobObject()

Dim device : Set device = job.CreateDeviceObject()
Dim pin : Set pin = job.CreatePinObject()
Dim sheet : Set sheet = job.CreateSheetObject()

Dim deviceIds, pinIds
Dim deviceCount, deviceIndex, pinCount, pinIndex
Dim deviceId, deviceName, pinId, pinName
Dim x, y, grid, col, row
Dim result
Dim hasUnplacedPins : hasUnplacedPins = False

deviceCount = job.GetTreeSelectedAllDeviceIds(deviceIds)

If deviceCount = 0 Then
    e3App.PutMessageEx 0, "Нет выбранных устройств.", 0, COLOR_RED, 0, 0
Else
    For deviceIndex = 1 To deviceCount
        deviceId = device.SetId(deviceIds(deviceIndex))
        deviceName = device.GetName()
        
        If device.GetPinIds(pinIds) = 0 Then
            e3App.PutMessageEx 0, "Устройство """ & deviceName & """ не содержит пинов.", _
                deviceIds(deviceIndex), COLOR_BLUE, 0, 0
        Else
            For pinIndex = 1 To UBound(pinIds)
                pinId = pin.SetId(pinIds(pinIndex))
                pinName = pin.GetName()
                
                result = pin.GetSchemaLocation(x, y, grid, col, row)
                
                If result = 0 Then
                    hasUnplacedPins = True
                    e3App.PutMessageEx 0, "Устройство """ & deviceName & """: пин """ & pinName & """ (ID: " & pinId & ") не размещён", _
                        deviceIds(deviceIndex), COLOR_RED, 0, 0
                End If
            Next
        End If
    Next

    If Not hasUnplacedPins Then
        e3App.PutMessageEx 0, "Все пины размещены.", 0, 0, COLOR_GREEN, 0
    End If
End If

' === Очистка объектов ===
Set sheet = Nothing
Set pin = Nothing
Set device = Nothing
Set job = Nothing
Set e3App = Nothing
