'*******************************************************************************
' Название скрипта: E3_UZ_CheckUnplacedSymbolsConcise
' Автор: E3.series VBScript Assistant
' Дата: 22.07.2025
' Описание: Скрипт для всех выделенных в дереве изделий находит их символы
'           и выводит информацию только о тех символах, которые не размещены
'           на схеме, с возможностью перехода к символу по ссылке.
'           Удалены все промежуточные сообщения о проверке/поиске.
'           Внимание: Отсутствуют проверки на успешное создание объектов.
'*******************************************************************************

Option Explicit

' Объявление переменных
Dim e3App
Dim job
Dim device
Dim symbol

Dim deviceIds
Dim deviceCount
Dim deviceId
Dim deviceName

Dim symbolIds
Dim symbolCount
Dim symbolId
Dim symbolName

Dim xMin, yMin, xMax, yMax
Dim result

' Определяем цвет для сообщений (красный, чтобы было заметно)
Const COLOR_RED = &HFF& ' Красный цвет (RGB)

On Error Resume Next ' Включаем обработку ошибок, чтобы избежать полного краха при отсутствии объектов

' Создаем объекты E3.series
Set e3App = CreateObject("CT.Application")
Set job = e3App.CreateJobObject()
Set device = job.CreateDeviceObject()
Set symbol = job.CreateSymbolObject()

e3App.PutInfo 0, "Начинаем поиск неразмещенных символов. Для перехода к символу нажмите на соответствующую строку."

' Получаем все выделенные изделия в дереве проекта
deviceCount = job.GetTreeSelectedAllDeviceIds(deviceIds)

If deviceCount > 0 Then
    For Each deviceId In deviceIds
        result = device.SetId(deviceId)
        If result <> 0 Then ' Проверяем, что установка ID изделия успешна
            deviceName = device.GetName()
            
            ' Получаем все символы для текущего изделия
            symbolCount = device.GetSymbolIds(symbolIds)
            
            If symbolCount > 0 Then
                For Each symbolId In symbolIds
                    result = symbol.SetId(symbolId)
                    If result <> 0 Then ' Проверяем, что установка ID символа успешна
                        symbolName = symbol.GetName()
                        
                        ' Проверяем, размещен ли символ на схеме
                        result = symbol.GetPlacedArea(xMin, yMin, xMax, yMax)
                        
                        If result = 0 Then ' Если GetPlacedArea вернул 0, символ НЕ РАЗМЕЩЕН
                            ' Выводим сообщение с ID символа в качестве LinkID, чтобы сделать его кликабельным
                            e3App.PutMessageEx 0, "НЕ РАЗМЕЩЕН: Символ '" & symbolName & "' (Изделие: " & deviceName & ")", symbolId, COLOR_RED, 0, 0
                        End If
                    End If ' Закрываем If для symbol.SetId
                Next
            End If ' Закрываем If для symbolCount > 0
        End If ' Закрываем If для device.SetId
    Next
Else
    e3App.PutInfo 0, "В дереве проекта не выделено ни одного изделия."
End If

e3App.PutInfo 0, "Поиск неразмещенных символов завершен."

' Освобождаем объекты
Set symbol = Nothing
Set device = Nothing
Set job = Nothing
Set e3App = Nothing

On Error GoTo 0