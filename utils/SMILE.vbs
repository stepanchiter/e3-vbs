' filepath: d:\E3_VBS_Scripts\utils\DrawSmile.vbs

Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject()

' Создаём объект листа
Set Sheet = Job.CreateSheetObject()

' Параметры для создания листа
Dim moduleId, sheetName, symbolName, position, isBeforePosition
moduleId = 0                        ' Без модуля
sheetName = "Smile"                ' Имя листа
symbolName = "Формат_А3_гор_1лист"               ' Формат A4 из базы данных Misc_Sheet
position = 0                       ' В конец проекта
isBeforePosition = 0              ' После указанной позиции

' Создаём новый лист
Dim sheetId, result, message
result = Sheet.Create(moduleId, sheetName, symbolName, position, isBeforePosition)

' Проверяем успешность создания
If result > 0 Then
    sheetId = result
    
    ' Создаем объект для рисования
    Set Graphic = Job.CreateGraphObject()
    
    ' Координаты центра смайлика
    Dim centerX, centerY, radius
    centerX = 100
    centerY = 100
    radius = 50
    If sheetId > 0 Then
        ' Рисуем лицо (окружность)
        result = Graphic.CreateCircle(sheetId, centerX, centerY, radius)
        ' Рисуем левый глаз
        result = Graphic.CreateCircle(sheetId, centerX - 20, centerY + 25, 5)
        ' Рисуем правый глаз
        result = Graphic.CreateCircle(sheetId, centerX + 20, centerY + 25, 5)
        ' Рисуем улыбку (дуга)
        result = Graphic.CreateArc(sheetId, centerX, centerY - 15, 25, 200, 340)

        If result = 0 Then
            message = "Ошибка при создании графики"
        Else
            message = "Смайлик успешно создан на листе"
        End If

        App.PutInfo 0, message
    End If
End If

Set Graphic = Nothing
Set Sheet = Nothing
Set Job = Nothing
Set App = Nothing