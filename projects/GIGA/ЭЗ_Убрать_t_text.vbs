' Инициализация констант для сообщений
Const INFO_MESSAGE = 0
Const WARNING_MESSAGE = 1
Const ERROR_MESSAGE = 2

' --- НАЧАЛО СКРИПТА ---

Set e3App = CreateObject("CT.Application")

If e3App Is Nothing Then
    WScript.Echo "Ошибка: Не удалось создать объект CT.Application. Убедитесь, что E3.series запущен."
    WScript.Quit
End If

' Инициализация объектов E3.series
Set job = e3App.CreateJobObject()

If job Is Nothing Then
    e3App.PutInfo ERROR_MESSAGE, "Ошибка: Не удалось создать объект Job."
    CleanupAndExit
End If

Set e3Text = job.CreateTextObject()

If e3Text Is Nothing Then
    e3App.PutInfo ERROR_MESSAGE, "Ошибка: Не удалось создать объект Text. Проверьте API E3.series."
    CleanupAndExit
End If

Dim selectedTextIds ' Массив для хранения ID выбранных текстовых объектов
Dim textCount ' Количество выбранных текстовых объектов
Dim processedTextCount : processedTextCount = 0 ' Счетчик измененных текстов

' Получаем ID всех выбранных текстовых объектов.
On Error Resume Next ' Включаем обработку ошибок на случай, если метод недоступен
textCount = job.GetSelectedTextIds(selectedTextIds)

If Err.Number <> 0 Then
    On Error GoTo 0
    e3App.PutInfo ERROR_MESSAGE, "Ошибка при вызове job.GetSelectedTextIds: " & Err.Description & ". Убедитесь, что текстовые объекты выбраны и метод доступен в вашей версии E3.series API."
    CleanupAndExit
End If
On Error GoTo 0 ' Отключаем обработку ошибок

If textCount = 0 Then
    e3App.PutInfo WARNING_MESSAGE, "Нет выбранных текстовых объектов для обработки. Пожалуйста, выберите надписи."
Else
    e3App.PutInfo INFO_MESSAGE, "Найдено " & textCount & " выбранных текстовых объектов. Начинаем обработку..."
    Dim currentTextId
    Dim currentText
    Dim newText
    Dim setResult

    For i = 1 To textCount
        currentTextId = selectedTextIds(i)
        
        e3Text.SetId currentTextId 
        
        currentText = e3Text.GetText() 

        If Left(currentText, 2) = "-t" Then
            newText = Replace(currentText, "-t", "-") 
            setResult = e3Text.SetText(newText) 

            If setResult = 0 Then
                e3App.PutInfo ERROR_MESSAGE, "Ошибка: Не удалось изменить текст '" & currentText & "' (ID: " & currentTextId & ") на '" & newText & "'."
            Else
                e3App.PutInfo INFO_MESSAGE, "Текст изменен: '" & currentText & "' -> '" & newText & "'"
                processedTextCount = processedTextCount + 1
            End If
        Else
            ' Если нет изменений, можно пропустить сообщение, чтобы не "засорять" лог,
            ' или оставить его для полной прозрачности.
            ' e3App.PutInfo INFO_MESSAGE, "Текст '" & currentText & "' (ID: " & currentTextId & ") не соответствует шаблону '-t'. Пропущен."
        End If
    Next
End If

If processedTextCount = 0 And textCount > 0 Then ' Если были выбранные тексты, но ни один не изменился
    e3App.PutInfo WARNING_MESSAGE, "Среди выбранных текстовых объектов не найдено ни одного, соответствующего шаблону '-t' для изменения."
ElseIf processedTextCount = 0 And textCount = 0 Then ' Если ничего не было выбрано изначально
    ' Сообщение уже было выведено в блоке If textCount = 0 Then
Else ' Если были изменения
    e3App.PutInfo INFO_MESSAGE, "Успешно изменено " & processedTextCount & " выбранных текстовых объектов."
End If

e3App.PutInfo INFO_MESSAGE, "Скрипт завершен."

' Функция для очистки объектов и выхода
Sub CleanupAndExit()
    Set e3Text = Nothing
    Set job = Nothing
    Set e3App = Nothing
    WScript.Quit
End Sub