'*******************************************************************************
' Название скрипта: E3_RoundSelectedDimensions
' Автор: E3.series VBScript Assistant
' Дата: 29.07.2025
' Описание: Скрипт для округления числовых значений размеров,
'           выделенных пользователем в E3.series, до ближайших 5 мм.
'           Предполагается, что значение размера всегда является числом.
'           Пользователь должен выделить размеры перед запуском скрипта.
'*******************************************************************************

Option Explicit

' --- Глобальные переменные ---
Dim e3Application
Dim job
Dim dimension
Dim message

' --- Константы ---
Const ROUND_TO_NEAREST = 5 ' Округление до 5 мм

' --- Инициализация объектов E3.series ---
Set e3Application = CreateObject("CT.Application")
If e3Application Is Nothing Then
    MsgBox "Не удалось подключиться к E3.series. Убедитесь, что E3.series запущен.", vbCritical
    WScript.Quit
End If

Set job = e3Application.CreateJobObject()
If job Is Nothing Then
    e3Application.PutInfo 2, "Ошибка: Не удалось создать объект Job."
    Set e3Application = Nothing
    WScript.Quit
End If

Set dimension = job.CreateDimensionObject()

' --- Получение выделенных размеров ---
Dim selectedDimensionIds()
Dim dimensionCount
dimensionCount = job.GetSelectedDimensionIds(selectedDimensionIds)

If dimensionCount = 0 Then
    e3Application.PutInfo 1, "Внимание: В проекте не выделено ни одного размера. Пожалуйста, выделите размеры, которые вы хотите округлить, и повторите попытку."
Else
    e3Application.PutInfo 0, "Найдено " & dimensionCount & " выделенных размера(ов)."

    Dim iDimensionIndex
    For iDimensionIndex = 1 To dimensionCount ' Массив в E3.series API часто начинается с 1
        Dim currentDimensionId
        currentDimensionId = selectedDimensionIds(iDimensionIndex)
        
        Dim result
        result = dimension.SetId(currentDimensionId)

        If result Then ' Успешно установили ID для объекта dimension
            Dim dimensionText
            Dim isTextUsed ' 0 - измерение, 1 - фиксированный текст
            Dim numericalValue
            Dim roundedValue
            
            result = dimension.GetText(dimensionText, isTextUsed)

            If result Then
                If IsNumeric(dimensionText) Then
                    numericalValue = CDbl(dimensionText)
                    roundedValue = RoundToNearest(numericalValue, ROUND_TO_NEAREST)

                    If roundedValue <> numericalValue Then
                        ' Устанавливаем округленное значение как фиксированный текст
                        ' Параметр '1' в SetText означает, что будет отображаться фиксированный текст.
                        result = dimension.SetText(CStr(roundedValue), 1) 
                        If result Then
                            e3Application.PutInfo 0, "Размер ID: " & currentDimensionId & " округлен с " & numericalValue & " до " & roundedValue
                        Else
                            e3Application.PutInfo 2, "Ошибка: Не удалось установить округленный текст для размера ID: " & currentDimensionId
                        End If
                    Else
                        e3Application.PutInfo 0, "Размер ID: " & currentDimensionId & " уже соответствует округленному значению (" & numericalValue & "). Пропущено."
                    End If
                Else
                    e3Application.PutInfo 1, "Внимание: Размер ID: " & currentDimensionId & " содержит нечисловое значение '" & dimensionText & "'. Пропущено."
                End If
            Else
                e3Application.PutInfo 2, "Ошибка: Не удалось получить текст размера для ID: " & currentDimensionId
            End If
        Else
            e3Application.PutInfo 2, "Ошибка: Не удалось установить ID для объекта Dimension (ID: " & currentDimensionId & "). Возможно, объект не существует или не является размером."
        End If
    Next
End If

e3Application.PutInfo 0, "Скрипт завершен."

' --- Очистка объектов ---
Set dimension = Nothing
Set job = Nothing
Set e3Application = Nothing

' --- Вспомогательная функция для округления ---
Function RoundToNearest(value, step)
    If step = 0 Then
        RoundToNearest = value
        Exit Function
    End If
    RoundToNearest = Round(value / step) * step
End Function