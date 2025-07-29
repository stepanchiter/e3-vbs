' Создаем объект приложения
Set e3Application = CreateObject("CT.Application")
Set job = e3Application.CreateJobObject()

' Получаем путь проекта
projectPath = job.GetPath()
If Len("" & projectPath) = 0 Then
    message = "Error getting the project path"
    e3Application.PutInfo 0, message
    WScript.Quit
Else
    message = "Project path is " & projectPath
    e3Application.PutInfo 0, message
End If

' Получаем имя проекта
projectName = job.GetName()
If Len("" & projectName) = 0 Then
    message = "Error getting the project name"
    e3Application.PutInfo 0, message
    WScript.Quit
Else
    message = "Name of the project is " & projectName
    e3Application.PutInfo 0, message
End If

' Формируем полный путь для сохранения e3d
Dim projectFile : projectFile = projectPath & "\" & projectName & ".e3d"
Dim useCompression : useCompression = False

' Сохраняем проект e3d
result = job.SaveAs(projectFile, useCompression)
Select Case result
    Case -1
        message = "Error saving project: No project open"
    Case -2 
        message = "Project could not be saved to " & projectFile
    Case -3 
        message = "Error saving project: Redliner intermediate project must have an e3n extension"         
    Case -4  
        message = "Error saving project: External document closed without saving the data"
    Case Else
        message = "Project is saved as " & projectFile
End Select 

e3Application.PutInfo 0, message

' Проверяем, содержит ли projectName "Sch2_" и заменяем на "Sch"
If InStr(projectName, "Sch2_") > 0 Then
    projectName = Replace(projectName, "Sch2_", "Sch")
    message = "Project name updated to " & projectName
    e3Application.PutInfo 0, message
End If

' Формируем полный путь для сохранения e3v
Dim projectFile1 : projectFile1 = projectPath & "\" & projectName & ".e3v"
Dim useCompression1 : useCompression1 = False

' Сохраняем проект e3v
result = job.SaveAs(projectFile1, useCompression1)
Select Case result
    Case -1
        message = "Error saving project: No project open"
    Case -2 
        message = "Project could not be saved to " & projectFile1
    Case -3 
        message = "Error saving project: Redliner intermediate project must have an e3n extension"         
    Case -4  
        message = "Error saving project: External document closed without saving the data"
    Case Else
        message = "Project is saved as " & projectFile1
End Select 

e3Application.PutInfo 0, message

' Освобождаем ресурсы
Set job = Nothing 
Set e3Application = Nothing