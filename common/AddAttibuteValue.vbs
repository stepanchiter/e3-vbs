'скрипт, который для каждого выделенного девайса находит имя его
'первого пина и создает атрибут (AttrFile5) для девайса. в значение атрибута
'записывает имя первого найденного пина.



Set e3Application = CreateObject("CT.Application") 
Set job = e3Application.CreateJobObject()
Set device = job.CreateDeviceObject()
Set pin = job.CreatePinObject()

' Настройки атрибута
attributeName = "AttrFile5"  ' Имя создаваемого атрибута

deviceCount = job.GetTreeSelectedAllDeviceIds(deviceIds) ' Получаем выделенные устройства

If deviceCount > 0 Then 
    For deviceIndex = 1 To deviceCount
        deviceId = device.SetId(deviceIds(deviceIndex))
        deviceName = device.GetName()
        
        ' Получаем все пины устройства
        result = device.GetAllPinIds(pinIds)    
        
        If result = 0 Then
            e3Application.PutInfo 0, "No pins found for device " & deviceName & " (" & deviceId & ")"
        Else        
            ' Берем первый пин
            pinId = pin.SetId(pinIds(1))
            pinName = pin.GetName()
            
            ' Добавляем атрибут
            addResult = device.AddAttibuteValue(attributeName, pinName)
            
            Select Case addResult
                Case 0
                    message = "Error adding attribute to device " & deviceName & " (" & deviceId & ")"
                Case -1
                    message = "Value too long for device " & deviceName & " (" & deviceId & ")"
                Case Else
                    message = "Device " & deviceName & " (" & deviceId & "): Attribute '" & attributeName & "' set to '" & pinName & "'"
            End Select
            
            e3Application.PutInfo 0, message
        End If    
    Next
Else
    e3Application.PutInfo 0, "No devices selected"
End If

Set pin = Nothing
Set device = Nothing
Set job = Nothing
Set e3Application = Nothing