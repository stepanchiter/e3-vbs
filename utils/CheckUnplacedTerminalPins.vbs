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

Dim terminalIds, pinIds
Dim terminalCount, terminalIndex, pinCount, pinIndex
Dim terminalId, terminalName, pinId, pinName
Dim x, y, grid, col, row
Dim result
Dim hasUnplacedPins : hasUnplacedPins = False

' Получаем все клеммники в проекте
terminalCount = job.GetTerminalIds(terminalIds)

If terminalCount = 0 Then
    e3App.PutMessageEx 0, "Клеммники не найдены!", 0, COLOR_RED, 0, 0
Else
    For terminalIndex = 1 To terminalCount
        terminalId = device.SetId(terminalIds(terminalIndex))
        terminalName = device.GetName()
        
        If device.GetPinIds(pinIds) <> 0 Then
            Dim firstUnplacedPinFound : firstUnplacedPinFound = False
            For pinIndex = 1 To UBound(pinIds)
                pinId = pin.SetId(pinIds(pinIndex))
                pinName = pin.GetName()
                
                result = pin.GetSchemaLocation(x, y, grid, col, row)
                
                If result = 0 And Not firstUnplacedPinFound Then
                    hasUnplacedPins = True
                    e3App.PutMessageEx 0, "Клеммник """ & terminalName & """: пин """ & pinName & """ (ID: " & pinId & ") не размещён", _
                        terminalIds(terminalIndex), COLOR_RED, 0, 0
                    firstUnplacedPinFound = True
                End If
            Next
        End If
    Next

    If Not hasUnplacedPins Then
        e3App.PutMessageEx 0, "Все пины всех клеммников размещены.", 0, 0, COLOR_GREEN, 0
    End If
End If

' === Очистка объектов ===
Set sheet = Nothing
Set pin = Nothing
Set device = Nothing
Set job = Nothing
Set e3App = Nothing
