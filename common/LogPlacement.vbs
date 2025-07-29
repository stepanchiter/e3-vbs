Set e3Application = CreateObject("CT.Application")  
Set job = e3Application.CreateJobObject()  
Set device = job.CreateDeviceObject()  

' Получаем все устройства в проекте
nDevices = job.GetAllDeviceIds(deviceIds)  

' Открываем файл для логов
Dim fso, logFile  
Set fso = CreateObject("Scripting.FileSystemObject")  
Set logFile = fso.OpenTextFile("C:\Scripts\placement_log.txt", 8, True)  

' Перебираем все устройства и логируем их
logFile.WriteLine "=== Размещено устройство: " & Now & " ==="

For i = 0 To nDevices - 1  
    device.SetId deviceIds(i)  
    logFile.WriteLine "Устройство: " & device.GetName() & " (ID: " & device.GetId() & ")"  
Next  

logFile.WriteLine "======================================="  
logFile.Close  

' Вывод в PutInfo
e3Application.PutInfo 0, "? Лог размещённых устройств обновлён!"  

' Освобождаем объекты  
Set logFile = Nothing  
Set fso = Nothing  
Set device = Nothing  
Set job = Nothing  
Set e3Application = Nothing  
