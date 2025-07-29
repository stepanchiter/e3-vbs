Set e3Application = CreateObject("CT.Application")
Set job = e3Application.CreateJobObject()
Set sheet = job.CreateSheetObject()

result = job.GetAllSheetIds(sheetIds)

If result > 0 Then

    For i = 1 To result
        sheetId = sheet.SetId(sheetIds(i))

        attrValue = sheet.GetAttributeValue("Тип документа")

        If attrValue = "Перечень элементов" Or attrValue = "Таблица соединений" Then
            sheetName = sheet.GetName()
            sheetAssignment = sheet.GetAssignment()
            sheetLocation = sheet.GetLocation()

            delResult = sheet.Delete()

            Select Case delResult
                Case 0
                    message = "Удалён лист: " & sheetName & " " & sheetAssignment & " " & sheetLocation & " (" & sheetId & ")"
                Case -1
                    message = "Ошибка удаления: не удалось заблокировать лист"
                Case -2
                    message = "Ошибка удаления: лист остаётся оффлайн после блокировки"
                Case -3
                    message = "Ошибка удаления: лист заблокирован"
                Case -4
                    message = "Ошибка удаления: лист является регионом"
                Case -5
                    message = "Ошибка удаления: лист только для чтения"
                Case -6
                    message = "Ошибка удаления: лист не существует"
            End Select

            e3Application.PutInfo 0, message
        End If
    Next

Else
    e3Application.PutInfo 0, "В проекте нет листов."
End If

Set sheet = Nothing
Set job = Nothing
Set e3Application = Nothing
