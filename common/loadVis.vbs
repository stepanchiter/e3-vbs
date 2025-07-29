Set e3Application = CreateObject("CT.Application")  
Set job = e3Application.CreateJobObject()  

' Указываем путь к файлу настроек уровней видимости
Dim levelFile  
levelFile = "C:\Templates\default_visibility.vis"  ' Укажите свой путь к файлу  

e3Application.PutInfo 0, "Загрузка уровней видимости из: " & levelFile  

' Загружаем уровни видимости
result = job.LoadLevelConfiguration(levelFile)  

If result = 0 Then  
    e3Application.PutInfo 0, "Ошибка загрузки уровней видимости из " & levelFile  
Else  
    e3Application.PutInfo 0, "Уровни видимости успешно загружены из " & levelFile  
End If  

' Освобождаем объекты
Set job = Nothing  
Set e3Application = Nothing  
