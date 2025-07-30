'*******************************************************************************
' Название скрипта: E3_FindDeviceAndPin_FromExcelInput_BRS_Approach
' Автор: E3.series VBScript Assistant
' Дата: 16.07.2025
' Описание: Скрипт для чтения имени устройства (F3) и пина (G3) из Excel.
'           Поиск выполняется в активном проекте E3.series.
'           Использует подход подключения к E3.series, аналогичный скриптам из БРС,
'           без явного вызова GetActiveJobId().
'           Изменено: Поиск устройства теперь осуществляется путем итерации по всем
'           устройствам проекта и сопоставления имени без учета регистра.
'           Также, поиск пина осуществляется путем итерации по всем пинам
'           найденного устройства и сопоставления имени без учета регистра.
'           Исправлено: Сообщение о завершении скрипта теперь выводится до очистки объектов E3.series.
'           Добавлено: Запись значений из ячеек Excel (A3-E3) в атрибуты найденного пина.
'           Добавлено: Обработка всех заполненных строк в таблице Excel, начиная со строки 3.
'           Добавлено: Путь к Excel файлу задается в настройках скрипта.
'*******************************************************************************

Option Explicit

' --- Настройки ---
Const EXCEL_SHEET_NAME = "Лист1"         ' Укажите имя листа в Excel, если оно отличается
Const START_DATA_ROW = 3               ' Начальная строка для чтения данных в Excel
' >>> Добавлена новая настройка для пути к файлу Excel <<<
Const EXCEL_FILE_PATH_DEFAULT = "D:\E3_VBS_Scripts\projects\GIGA\EXEL\запись атрибутов в модули.xlsx" ' Ваш путь к файлу Excel по умолчанию

' --- Главная подпрограмма ---
Call Main()

Sub Main()
    ' --- Инициализация объектов E3.series ---
    Dim e3App, job, device, pin
    Dim deviceName, pinName
    Dim deviceId, pinId
    Dim EXCEL_FILE_PATH ' Теперь это переменная, а не константа, чтобы можно было перезаписать
    
    ' Переменные для чтения атрибутов из Excel
    Dim tagPosition, tagDescription, plcSignalType, plcConnectionType, plcUnit
    
    On Error Resume Next
    ' Попытка получить уже запущенный экземпляр E3.series
    Set e3App = GetObject(, "CT.Application")
    
    If e3App Is Nothing Then
        ' Если E3.series не запущен, пытаемся создать новый экземпляр
        Set e3App = CreateObject("CT.Application")
        If e3App Is Nothing Then
            MsgBox "E3.series Application не запущен или не найден.", vbCritical, "Ошибка E3.series"
            Exit Sub ' Выход из подпрограммы
        End If
    End If
    On Error GoTo 0 ' Выключаем обработку ошибок после инициализации e3App
    
    ' Создаем объект job. В этом подходе предполагается, что он будет работать
    ' с активным проектом, если он открыт.
    Set job = e3App.CreateJobObject()
    
    ' Проверка, что job объект успешно создан и проект открыт (косвенно)
    ' Если job.CreateDeviceObject() не сработает, это укажет на проблему с проектом.
    On Error Resume Next
    Set device = job.CreateDeviceObject()
    If device Is Nothing Then
        e3App.PutInfo 2, "Не удалось создать объект Device. Убедитесь, что проект E3.series открыт."
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    On Error GoTo 0
    
    Set pin = job.CreatePinObject()
    
    e3App.PutInfo 0, "Скрипт запущен: Поиск устройства и пина из Excel (подход БРС)."
    
    ' --- Определение пути к Excel файлу ---
    EXCEL_FILE_PATH = EXCEL_FILE_PATH_DEFAULT ' По умолчанию используем путь из настроек
    
    ' Проверяем, существует ли файл по пути по умолчанию
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(EXCEL_FILE_PATH) Then
        e3App.PutInfo 1, "Файл Excel по пути по умолчанию '" & EXCEL_FILE_PATH & "' не найден."
        ' Если файла нет, запрашиваем путь у пользователя
        EXCEL_FILE_PATH = InputBox("Файл Excel по пути по умолчанию не найден. Пожалуйста, введите полный путь к вашему Excel файлу:", "Путь к Excel файлу", "C:\Temp\ВашФайл.xlsx")
    End If
    Set fso = Nothing

    If Trim(EXCEL_FILE_PATH) = "" Then
        e3App.PutInfo 2, "Путь к Excel файлу не был введен. Скрипт отменен."
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' Выход из подпрограммы
    End If
    
    ' --- Инициализация объектов Excel ---
    Dim objExcel, objWorkbook, objSheet
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If objExcel Is Nothing Then
        e3App.PutInfo 2, "Не удалось запустить приложение Excel. Убедитесь, что Excel установлен."
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' Выход из подпрограммы
    End If
    objExcel.Visible = False ' Скрыть Excel
    
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_FILE_PATH)
    If objWorkbook Is Nothing Then
        e3App.PutInfo 2, "Не удалось открыть файл Excel: " & EXCEL_FILE_PATH & ". Проверьте, что файл существует и не занят."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' Выход из подпрограммы
    End If
    
    Set objSheet = objWorkbook.Sheets(EXCEL_SHEET_NAME)
    If objSheet Is Nothing Then
        e3App.PutInfo 2, "Не удалось найти лист '" & EXCEL_SHEET_NAME & "' в файле Excel. Проверьте имя листа."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' Выход из подпрограммы
    End If
    On Error GoTo 0 ' Отключаем обработку ошибок после успешного открытия Excel
    
    ' --- Определяем последнюю заполненную строку в листе Excel ---
    Dim lastRow
    On Error Resume Next
    ' Использование SpecialCells(xlCellTypeLastCell) для нахождения последней ячейки с данными
    lastRow = objSheet.Cells.SpecialCells(11).Row ' xlCellTypeLastCell = 11
    If Err.Number <> 0 Then
        e3App.PutInfo 1, "Не удалось определить последнюю строку в Excel. Возможно, лист пуст. Ошибка: " & Err.Description
        lastRow = START_DATA_ROW ' В случае ошибки, устанавливаем на начальную строку
        Err.Clear
    End If
    On Error GoTo 0

    If lastRow < START_DATA_ROW Then
        e3App.PutInfo 1, "В Excel файле нет данных, начиная со строки " & START_DATA_ROW & ". Скрипт завершен."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub
    End If

    ' --- Запускаем цикл по всем заполненным строкам ---
    Dim currentRow
    For currentRow = START_DATA_ROW To lastRow
        e3App.PutInfo 0, "--- Обработка строки: " & currentRow & " ---"
        
        ' --- Чтение имени устройства и пина из Excel для текущей строки ---
        deviceName = Trim(CStr(objSheet.Cells(currentRow, 6).Value)) ' Столбец F
        pinName = Trim(CStr(objSheet.Cells(currentRow, 7).Value))    ' Столбец G
        
        ' --- Чтение значений для атрибутов из Excel для текущей строки ---
        tagPosition = Trim(CStr(objSheet.Cells(currentRow, 1).Value))  ' Столбец A
        tagDescription = Trim(CStr(objSheet.Cells(currentRow, 2).Value)) ' Столбец B
        plcSignalType = Trim(CStr(objSheet.Cells(currentRow, 3).Value)) ' Столбец C
        plcConnectionType = Trim(CStr(objSheet.Cells(currentRow, 4).Value)) ' Столбец D
        plcUnit = Trim(CStr(objSheet.Cells(currentRow, 5).Value))      ' Столбец E
        
        ' --- Проверка прочитанных значений и выполнение поиска ---
        If deviceName = "" Or pinName = "" Then
            e3App.PutInfo 1, "Строка " & currentRow & ": Пропущена, так как имя устройства (F" & currentRow & ") или имя пина (G" & currentRow & ") пусты."
            ' Продолжаем цикл к следующей строке
        Else
            e3App.PutInfo 0, "Строка " & currentRow & " Прочитано из Excel: Устройство = '" & deviceName & "', Пин = '" & pinName & "'"
            e3App.PutInfo 0, "Значения атрибутов из Excel: TAG Позиция='" & tagPosition & "', TAG Описание='" & tagDescription & "', ПЛК-Тип сигнала='" & plcSignalType & "', ПЛК-Тип подключения='" & plcConnectionType & "', ПЛК-Единица измерения='" & plcUnit & "'"
            
            ' --- Поиск устройства: теперь итерация по всем устройствам ---
            e3App.PutInfo 0, "Ищу устройство: '" & deviceName & "' путем полного перебора..."
            Dim allDeviceIds, totalDeviceCount
            Dim currentDeviceId
            Dim foundDeviceId : foundDeviceId = 0 ' Переменная для хранения ID найденного устройства

            totalDeviceCount = job.GetAllDeviceIds(allDeviceIds) ' Получаем все ID устройств в проекте

            If totalDeviceCount > 0 Then
                Dim k ' Используем другую переменную для цикла, чтобы избежать конфликтов с 'i' в родительском скрипте
                For k = 1 To totalDeviceCount
                    currentDeviceId = allDeviceIds(k)
                    device.SetId currentDeviceId ' Инициализируем объект устройства текущим ID
                    Dim currentDeviceName
                    currentDeviceName = device.GetName() ' Получаем имя текущего устройства

                    ' Сравниваем имена без учета регистра
                    If LCase(currentDeviceName) = LCase(deviceName) Then
                        foundDeviceId = currentDeviceId ' Устройство найдено, сохраняем ID
                        Exit For ' Выходим из цикла, так как нашли совпадение
                    End If
                Next
            End If

            If foundDeviceId = 0 Then
                e3App.PutInfo 1, "Строка " & currentRow & ": Устройство '" & deviceName & "' не найдено в проекте."
            Else
                deviceId = foundDeviceId ' Присваиваем найденный ID для дальнейшего использования
                device.SetId deviceId ' Инициализируем объект устройства найденным ID
                e3App.PutInfo 0, "Строка " & currentRow & ": Устройство '" & deviceName & "' найдено. ID: " & deviceId
                
                ' --- Поиск пина на найденном устройстве: теперь итерация по всем пинам устройства ---
                e3App.PutInfo 0, "Ищу пин: '" & pinName & "' на устройстве '" & deviceName & "' путем полного перебора..."
                Dim allPinIds, totalPinCount
                Dim currentPinId
                Dim foundPinId : foundPinId = 0 ' Переменная для хранения ID найденного пина

                totalPinCount = device.GetAllPinIds(allPinIds) ' Получаем все ID пинов на найденном устройстве

                If totalPinCount > 0 Then
                    Dim l ' Используем другую переменную для цикла
                    For l = 1 To totalPinCount
                        currentPinId = allPinIds(l)
                        pin.SetId currentPinId ' Инициализируем объект пина текущим ID
                        Dim currentPinName
                        currentPinName = pin.GetName() ' Получаем имя текущего пина

                        ' Сравниваем имена пинов без учета регистра
                        If LCase(currentPinName) = LCase(pinName) Then
                            foundPinId = currentPinId ' Пин найден, сохраняем ID
                            Exit For ' Выходим из цикла, так как нашли совпадение
                        End If
                    Next
                End If

                If foundPinId = 0 Then
                    e3App.PutInfo 1, "Строка " & currentRow & ": Пин '" & pinName & "' не найден на устройстве '" & deviceName & "'."
                Else
                    pinId = foundPinId ' Присваиваем найденный ID для дальнейшего использования
                    pin.SetId pinId ' Инициализируем объект пина найденным ID
                    e3App.PutInfo 0, "Строка " & currentRow & ": Пин '" & pinName & "' найден на устройстве '" & deviceName & "'. ID пина: " & pinId
                    
                    ' --- Запись значений в атрибуты найденного пина ---
                    e3App.PutInfo 0, "Строка " & currentRow & ": Запись атрибутов для пина '" & pinName & "'..."
                    
                    On Error Resume Next ' Включаем обработку ошибок для SetAttributeValue
                    
                    ' Проверяем и записываем 'TAG Позиция'
                    If tagPosition <> "" Then
                        If pin.SetAttributeValue("TAG Позиция", tagPosition) = 0 Then
                            e3App.PutInfo 1, "Строка " & currentRow & ": Ошибка при установке атрибута 'TAG Позиция' для пина '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "Строка " & currentRow & ": Атрибут 'TAG Позиция' успешно установлен в '" & tagPosition & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "Строка " & currentRow & ": Значение для 'TAG Позиция' в Excel пусто, атрибут не изменен."
                    End If
                    
                    ' Проверяем и записываем 'TAG Описание'
                    If tagDescription <> "" Then
                        If pin.SetAttributeValue("TAG Описание", tagDescription) = 0 Then
                            e3App.PutInfo 1, "Строка " & currentRow & ": Ошибка при установке атрибута 'TAG Описание' для пина '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "Строка " & currentRow & ": Атрибут 'TAG Описание' успешно установлен в '" & tagDescription & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "Строка " & currentRow & ": Значение для 'TAG Описание' в Excel пусто, атрибут не изменен."
                    End If

                    ' Проверяем и записываем 'ПЛК - Тип сигнала'
                    If plcSignalType <> "" Then
                        If pin.SetAttributeValue("ПЛК - Тип сигнала", plcSignalType) = 0 Then
                            e3App.PutInfo 1, "Строка " & currentRow & ": Ошибка при установке атрибута 'ПЛК - Тип сигнала' для пина '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "Строка " & currentRow & ": Атрибут 'ПЛК - Тип сигнала' успешно установлен в '" & plcSignalType & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "Строка " & currentRow & ": Значение для 'ПЛК - Тип сигнала' в Excel пусто, атрибут не изменен."
                    End If

                    ' Проверяем и записываем 'ПЛК - Тип подключения'
                    If plcConnectionType <> "" Then
                        If pin.SetAttributeValue("ПЛК - Тип подключения", plcConnectionType) = 0 Then
                            e3App.PutInfo 1, "Строка " & currentRow & ": Ошибка при установке атрибута 'ПЛК - Тип подключения' для пина '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "Строка " & currentRow & ": Атрибут 'ПЛК - Тип подключения' успешно установлен в '" & plcConnectionType & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "Строка " & currentRow & ": Значение для 'ПЛК - Тип подключения' в Excel пусто, атрибут не изменен."
                    End If

                    ' Проверяем и записываем 'ПЛК - Единица измерения'
                    If plcUnit <> "" Then
                        If pin.SetAttributeValue("ПЛК - Единица измерения", plcUnit) = 0 Then
                            e3App.PutInfo 1, "Строка " & currentRow & ": Ошибка при установке атрибута 'ПЛК - Единица измерения' для пина '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "Строка " & currentRow & ": Атрибут 'ПЛК - Единица измерения' успешно установлен в '" & plcUnit & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "Строка " & currentRow & ": Значение для 'ПЛК - Единица измерения' в Excel пусто, атрибут не изменен."
                    End If

                    On Error GoTo 0 ' Выключаем обработку ошибок после установки атрибутов
                End If ' End If foundPinId = 0
            End If ' End If foundDeviceId = 0
        End If ' End If deviceName = "" Or pinName = ""
    Next ' Next currentRow

    ' --- Обязательная очистка Excel объектов ---
    Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
    
    ' --- Выводим сообщение о завершении скрипта ПЕРЕД очисткой объектов E3.series ---
    e3App.PutInfo 0, "Скрипт завершен."

    ' --- Очистка объектов E3.series ---
    Call CleanUpE3Objects(pin, device, job, e3App)
    
End Sub ' End Sub Main()

' --- Вспомогательные подпрограммы для очистки объектов ---
Sub CleanUpExcelObjects(ByRef objSheet, ByRef objWorkbook, ByRef objExcel)
    On Error Resume Next ' Включаем обработку ошибок для очистки
    If Not objWorkbook Is Nothing Then objWorkbook.Close False ' Закрываем книгу без сохранения
    If Not objExcel Is Nothing Then objExcel.Quit ' Закрываем приложение Excel
    Set objSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0 ' Выключаем обработку ошибок
End Sub

Sub CleanUpE3Objects(ByRef pin, ByRef device, ByRef job, ByRef e3App)
    On Error Resume Next ' Включаем обработку ошибок для очистки
    Set pin = Nothing
    Set device = Nothing
    Set job = Nothing
    Set e3App = Nothing
    On Error GoTo 0 ' Выключаем обработку ошибок
End Sub