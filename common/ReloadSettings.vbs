Set e3Application = CreateObject("CT.Application")  
Set job = e3Application.CreateJobObject()  

' Указываем путь к файлу шаблона настроек
Dim templateFile  
templateFile = "D:\E3_config_backup\default.e3t"  ' Укажите свой путь к файлу шаблона  

e3Application.PutInfo 0, "Перезагрузка настроек проекта из: " & templateFile  

' Вызываем метод ReloadSettings()
result = job.ReloadSettings(templateFile)  

If result = True Then  
    e3Application.PutInfo 0, "? Настройки успешно загружены"
Else  
    e3Application.PutInfo 0, "? Загрузка настроек с ошибками"
End If  

' Освобождаем объекты
Set job = Nothing  
Set e3Application = Nothing  
