' filepath: d:\E3_VBS_Scripts\utils\DrawSmile.vbs

Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject()

' Создаём новый лист
Set Sheet = Job.CreateSheetObject()
Sheet.SetName "Smile"
Sheet.SetFormat "A4"
' Sheet.SetOrientation 1 ' 1 = альбомная, 0 = книжная

' Координаты центра смайлика
Dim centerX, centerY, radius
centerX = 100
centerY = 100
radius = 50

Set Graphic = Job.CreateGraphObject()

Dim sheetCount, sheetIds, sheetId, result, message
sheetCount = Job.GetTreeSelectedSheetIds(sheetIds)

If sheetCount > 0 Then
    sheetId = Sheet.SetId(sheetIds(1))
    If sheetId > 0 Then
        ' Рисуем лицо (окружность)
        result = Graphic.CreateCircle(sheetId, centerX, centerY, radius)
        ' Рисуем левый глаз
        result = Graphic.CreateCircle(sheetId, centerX - 20, centerY - 15, 5)
        ' Рисуем правый глаз
        result = Graphic.CreateCircle(sheetId, centerX + 20, centerY - 15, 5)
        ' Рисуем улыбку (дуга)
        result = Graphic.CreateArc(sheetId, centerX, centerY + 10, 25, 200, 340)

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