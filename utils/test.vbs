'*******************************************************************************
' Название скрипта: E3_FindPinLocation_FromExcelInput
' Автор: E3.series VBScript Assistant
' Дата: 06.08.2025
' Описание: Скрипт для чтения имени устройства (F3) и пина (G3) из Excel.
'           Поиск выполняется в активном проекте E3.series.
'           Для каждого найденного пина определяется его позиция на схеме.
'           Обработка всех заполненных строк в таблице Excel, начиная со строки 3.
'*******************************************************************************

Option Explicit

' --- Настройки ---
Const EXCEL_SHEET_NAME = "Лист1"         ' Укажите имя листа в Excel, если оно отличается
Const START_DATA_ROW = 3               ' Начальная строка для чтения данных в Excel
' >>> Путь к файлу Excel по умолчанию <<<
Const EXCEL_FILE_PATH_DEFAULT = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\ЗИС\ПРИНЦИП\ЩУ2\запись атрибутов в модули.xlsx"
' >>> Настройки для вставки фрагмента <<<
Const FRAGMENT_PATH = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\terminal2.e3p"
Const Y_OFFSET = 150                    ' Смещение по Y от координат пина (отнимаем от Y)

' --- Главная подпрограмма ---
Call Main()

Sub Main()
    ' --- Инициализация объектов E3.series ---
    Dim e3App, job, device, pin, sheet
    Dim deviceName, pinName
    Dim deviceId, pinId
    Dim EXCEL_FILE_PATH
    
    On Error Resume Next
    ' Попытка получить уже запущенный экземпляр E3.series
    Set e3App = GetObject(, "CT.Application")
    
    If e3App Is Nothing Then
        ' Если E3.series не запущен, пытаемся создать новый экземпляр
        Set e3App = CreateObject("CT.Application")
        If e3App Is Nothing Then
            MsgBox "E3.series Application не запущен или не найден.", vbCritical, "Ошибка E3.series"
            Exit Sub
        End If
    End If
    On Error GoTo 0
    
    ' Создаем объекты E3.series
    Set job = e3App.CreateJobObject()
    
    ' Проверка, что job объект успешно создан
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
    Set sheet = job.CreateSheetObject() ' Добавляем объект для работы с листами схемы
    
    e3App.PutInfo 0, "Скрипт запущен: Поиск позиций пинов на схеме из Excel."
    
    ' --- Определение пути к Excel файлу ---
    EXCEL_FILE_PATH = EXCEL_FILE_PATH_DEFAULT
    
    ' Проверяем, существует ли файл по пути по умолчанию
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(EXCEL_FILE_PATH) Then
        e3App.PutInfo 1, "Файл Excel по пути по умолчанию '" & EXCEL_FILE_PATH & "' не найден."
        EXCEL_FILE_PATH = InputBox("Файл Excel по пути по умолчанию не найден. Пожалуйста, введите полный путь к вашему Excel файлу:", "Путь к Excel файлу", "C:\Temp\ВашФайл.xlsx")
    End If
    Set fso = Nothing

    If Trim(EXCEL_FILE_PATH) = "" Then
        e3App.PutInfo 2, "Путь к Excel файлу не был введен. Скрипт отменен."
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    
    ' --- Инициализация объектов Excel ---
    Dim objExcel, objWorkbook, objSheet
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If objExcel Is Nothing Then
        e3App.PutInfo 2, "Не удалось запустить приложение Excel. Убедитесь, что Excel установлен."
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    objExcel.Visible = False
    
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_FILE_PATH)
    If objWorkbook Is Nothing Then
        e3App.PutInfo 2, "Не удалось открыть файл Excel: " & EXCEL_FILE_PATH & ". Проверьте, что файл существует и не занят."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    
    Set objSheet = objWorkbook.Sheets(EXCEL_SHEET_NAME)
    If objSheet Is Nothing Then
        e3App.PutInfo 2, "Не удалось найти лист '" & EXCEL_SHEET_NAME & "' в файле Excel. Проверьте имя листа."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    On Error GoTo 0
    
    ' --- Определяем последнюю заполненную строку в листе Excel ---
    Dim lastRow
    On Error Resume Next
    lastRow = objSheet.Cells.SpecialCells(11).Row ' xlCellTypeLastCell = 11
    If Err.Number <> 0 Then
        e3App.PutInfo 1, "Не удалось определить последнюю строку в Excel. Возможно, лист пуст. Ошибка: " & Err.Description
        lastRow = START_DATA_ROW
        Err.Clear
    End If
    On Error GoTo 0

    If lastRow < START_DATA_ROW Then
        e3App.PutInfo 1, "В Excel файле нет данных, начиная со строки " & START_DATA_ROW & ". Скрипт завершен."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If

    ' --- Запускаем цикл по всем заполненным строкам ---
    Dim currentRow
    For currentRow = START_DATA_ROW To lastRow
        e3App.PutInfo 0, "--- Обработка строки: " & currentRow & " ---"
        
        ' --- Чтение имени устройства и пина из Excel для текущей строки ---
        deviceName = Trim(CStr(objSheet.Cells(currentRow, 6).Value)) ' Столбец F
        pinName = Trim(CStr(objSheet.Cells(currentRow, 7).Value))    ' Столбец G
        
        ' --- Чтение условия для вставки фрагмента из столбца H ---
        Dim fragmentCondition
        fragmentCondition = Trim(CStr(objSheet.Cells(currentRow, 8).Value)) ' Столбец H
        
        ' --- Чтение значений для мастерпинов из столбцов I и J ---
        Dim masterPin1, masterPin2
        masterPin1 = Trim(CStr(objSheet.Cells(currentRow, 9).Value))  ' Столбец I - для первого устройства
        masterPin2 = Trim(CStr(objSheet.Cells(currentRow, 10).Value)) ' Столбец J - для второго устройства
        
        ' --- Проверяем условие "-XT" сразу ---
        If InStr(1, fragmentCondition, "-XT", vbTextCompare) = 0 Then
            e3App.PutInfo 0, "Строка " & currentRow & ": Условие '-XT' не найдено в столбце H ('" & fragmentCondition & "'). Строка пропущена."
            ' Переходим к следующей строке без дальнейшей обработки
        ElseIf deviceName = "" Or pinName = "" Then
            e3App.PutInfo 1, "Строка " & currentRow & ": Найдено условие '-XT', но имя устройства (F" & currentRow & ") или имя пина (G" & currentRow & ") пусты. Строка пропущена."
        Else
            e3App.PutInfo 0, "Строка " & currentRow & " Найдено условие '-XT'. Обработка: Устройство = '" & deviceName & "', Пин = '" & pinName & "', MasterPin1 = '" & masterPin1 & "', MasterPin2 = '" & masterPin2 & "'"
            
            ' --- Поиск устройства: итерация по всем устройствам ---
            e3App.PutInfo 0, "Ищу устройство: '" & deviceName & "' путем полного перебора..."
            Dim allDeviceIds, totalDeviceCount
            Dim currentDeviceId
            Dim foundDeviceId : foundDeviceId = 0

            totalDeviceCount = job.GetAllDeviceIds(allDeviceIds)

            If totalDeviceCount > 0 Then
                Dim k
                For k = 1 To totalDeviceCount
                    currentDeviceId = allDeviceIds(k)
                    device.SetId currentDeviceId
                    Dim currentDeviceName
                    currentDeviceName = device.GetName()

                    If LCase(currentDeviceName) = LCase(deviceName) Then
                        foundDeviceId = currentDeviceId
                        Exit For
                    End If
                Next
            End If

            If foundDeviceId = 0 Then
                e3App.PutInfo 1, "Строка " & currentRow & ": Устройство '" & deviceName & "' не найдено в проекте."
            Else
                deviceId = foundDeviceId
                device.SetId deviceId
                e3App.PutInfo 0, "Строка " & currentRow & ": Устройство '" & deviceName & "' найдено. ID: " & deviceId
                
                ' --- Поиск пина на найденном устройстве ---
                e3App.PutInfo 0, "Ищу пин: '" & pinName & "' на устройстве '" & deviceName & "' путем полного перебора..."
                Dim allPinIds, totalPinCount
                Dim currentPinId
                Dim foundPinId : foundPinId = 0

                totalPinCount = device.GetAllPinIds(allPinIds)

                If totalPinCount > 0 Then
                    Dim l
                    For l = 1 To totalPinCount
                        currentPinId = allPinIds(l)
                        pin.SetId currentPinId
                        Dim currentPinName
                        currentPinName = pin.GetName()

                        If LCase(currentPinName) = LCase(pinName) Then
                            foundPinId = currentPinId
                            Exit For
                        End If
                    Next
                End If

                If foundPinId = 0 Then
                    e3App.PutInfo 1, "Строка " & currentRow & ": Пин '" & pinName & "' не найден на устройстве '" & deviceName & "'."
                Else
                    pinId = foundPinId
                    pin.SetId pinId
                    e3App.PutInfo 0, "Строка " & currentRow & ": Пин '" & pinName & "' найден на устройстве '" & deviceName & "'. ID пина: " & pinId
                    
                    ' --- Получение позиции пина на схеме ---
                    e3App.PutInfo 0, "Строка " & currentRow & ": Определение позиции пина '" & pinName & "' на схеме..."
                    
                    Dim xPosition, yPosition, gridDescription, columnValue, rowValue
                    Dim result
                    
                    On Error Resume Next
                    result = pin.GetSchemaLocation(xPosition, yPosition, gridDescription, columnValue, rowValue)
                    On Error GoTo 0
                    
                    If result = 0 Then
                        e3App.PutInfo 1, "Строка " & currentRow & ": Пин '" & pinName & "' (" & pinId & ") не размещен на схеме или произошла ошибка."
                    Else
                        ' Устанавливаем ID листа схемы для получения его имени
                        sheet.SetId result
                        Dim sheetName
                        sheetName = sheet.GetName()
                        
                        e3App.PutInfo 0, "Строка " & currentRow & ": Позиция пина '" & pinName & "' (" & pinId & ") на схеме:"
                        e3App.PutInfo 0, "    Лист схемы: " & sheetName & " (ID: " & result & ")"
                        e3App.PutInfo 0, "    X координата: " & xPosition
                        e3App.PutInfo 0, "    Y координата: " & yPosition
                        e3App.PutInfo 0, "    Сетка: " & gridDescription
                        e3App.PutInfo 0, "    Столбец: " & columnValue
                        e3App.PutInfo 0, "    Строка: " & rowValue
                        
                        ' --- Вставка фрагмента по координатам пина (условие "-XT" уже проверено) ---
                        Dim fragmentX, fragmentY
                        fragmentX = xPosition
                        fragmentY = yPosition - Y_OFFSET ' Отнимаем смещение от Y координаты
                        
                        e3App.PutInfo 0, "Строка " & currentRow & ": Вставка фрагмента terminal2.e3p по координатам X=" & fragmentX & ", Y=" & fragmentY
                        
                        Dim fragmentResult
                        fragmentResult = PlaceFragmentOnSheet(sheet, result, FRAGMENT_PATH, "", fragmentX, fragmentY, e3App, currentRow)
                        
                        ' --- Переименование устройств -sXT1 после успешной вставки фрагмента ---
                        If fragmentResult = 0 Then
                            e3App.PutInfo 0, "Строка " & currentRow & ": Фрагмент успешно вставлен. Запуск переименования устройств -sXT1..."
                            Call RenameDevicesAfterFragment(job, e3App, currentRow, masterPin1, masterPin2)
                        Else
                            e3App.PutInfo 1, "Строка " & currentRow & ": Фрагмент не вставлен (код: " & fragmentResult & "). Переименование не выполняется."
                        End If
                    End If
                End If
            End If
        End If
    Next

    ' --- Обязательная очистка Excel объектов ---
    Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
    
    ' --- Выводим сообщение о завершении скрипта ---
    e3App.PutInfo 0, "Скрипт завершен."

    ' --- Очистка объектов E3.series ---
    Call CleanUpE3Objects(pin, device, job, sheet, e3App)
    
End Sub

' --- Функция для вставки фрагмента на лист схемы ---
Function PlaceFragmentOnSheet(sheet, sheetId, fragmentPath, version, xPosition, yPosition, e3App, rowNumber)
    On Error Resume Next
    
    ' Устанавливаем ID целевого листа
    sheet.SetId sheetId
    Dim sheetName
    sheetName = sheet.GetName()
    
    ' Вставляем фрагмент
    Dim result
    result = sheet.PlacePart(fragmentPath, version, xPosition, yPosition, 0.0)
    
    ' Обрабатываем результат вставки
    Dim message
    Select Case result
        Case 9
            message = "Ошибка вставки фрагмента на лист " & sheetName & " (" & sheetId & "): Несовместимая версия файла фрагмента"
        Case 3
            message = "Ошибка вставки фрагмента на лист " & sheetName & " (" & sheetId & "): Неверное имя фрагмента или версия"
        Case 0
            message = "Фрагмент успешно вставлен на лист " & sheetName & " (" & sheetId & ") по координатам X=" & xPosition & ", Y=" & yPosition
        Case -1
            message = "Ошибка вставки фрагмента на лист " & sheetName & " (" & sheetId & "): Фрагмент состоит из нескольких листов и установлена настройка 'Игнорировать границы листа'"
        Case -2
            message = "Ошибка вставки фрагмента на лист " & sheetName & " (" & sheetId & "): Фрагмент содержит листы и не установлена настройка 'Игнорировать границы листа'"
        Case -3
            message = "Ошибка вставки фрагмента на лист " & sheetName & " (" & sheetId & "): Фрагмент уже размещен или другие объекты размещены по координатам X=" & xPosition & ", Y=" & yPosition
        Case -4
            message = "Ошибка вставки фрагмента на лист " & sheetName & " (" & sheetId & "): Лист заблокирован"
        Case Else
            message = "Ошибка вставки фрагмента на лист " & sheetName & " (" & sheetId & "): Код ошибки " & result
    End Select
    
    e3App.PutInfo 0, "Строка " & rowNumber & ": " & message
    
    On Error GoTo 0
    PlaceFragmentOnSheet = result
End Function

' --- Функция переименования устройств -sXT1 в -XT666 после вставки фрагмента ---
Sub RenameDevicesAfterFragment(job, e3App, rowNumber, masterPin1Value, masterPin2Value)
    On Error Resume Next
    
    Dim renameDevice
    Set renameDevice = job.CreateDeviceObject()
    
    Dim deviceIds
    Dim result
    result = job.GetAllDeviceIds(deviceIds)
    
    Dim foundCount
    foundCount = 0
    
    e3App.PutInfo 0, "Строка " & rowNumber & ": === ПОИСК УСТРОЙСТВ -sXT1 ==="
    
    If result > 0 Then
        e3App.PutInfo 0, "Строка " & rowNumber & ": Всего устройств в проекте: " & result
        
        ' Сначала найдем все устройства -sXT1 и проверим их мастерпины
        Dim sxtDevices()
        Dim sxtCount
        sxtCount = 0
        
        Dim i, name, currentMasterPin
        For i = 1 To result
            renameDevice.SetId deviceIds(i)
            name = renameDevice.GetName()
            
            If name = "-sXT1" Then
                ' Проверяем, есть ли у устройства мастерпин
                currentMasterPin = renameDevice.GetMasterPinName()
                
                e3App.PutInfo 0, "Строка " & rowNumber & ": --- Найдено устройство -sXT1 ---"
                e3App.PutInfo 0, "Строка " & rowNumber & ": ID устройства: " & deviceIds(i)
                e3App.PutInfo 0, "Строка " & rowNumber & ": Текущий мастерпин: '" & currentMasterPin & "'"
                
                ' Если мастерпин существует (не пустой)
                If Len(Trim(currentMasterPin)) > 0 Then
                    sxtCount = sxtCount + 1
                    ReDim Preserve sxtDevices(sxtCount - 1)
                    sxtDevices(sxtCount - 1) = deviceIds(i)
                    e3App.PutInfo 0, "Строка " & rowNumber & ": Устройство добавлено в список для обработки (№" & sxtCount & ")"
                Else
                    e3App.PutInfo 0, "Строка " & rowNumber & ": Устройство пропущено - нет мастерпина"
                End If
            End If
        Next
        
        e3App.PutInfo 0, "Строка " & rowNumber & ": === ОБРАБОТКА НАЙДЕННЫХ УСТРОЙСТВ ==="
        e3App.PutInfo 0, "Строка " & rowNumber & ": Устройств с мастерпинами для обработки: " & sxtCount
        
        ' Теперь обрабатываем найденные устройства с мастерпинами
        Dim j, newName, resultSet, resultPin, finalMasterPin
        For j = 0 To sxtCount - 1
            If j >= 2 Then
                e3App.PutInfo 0, "Строка " & rowNumber & ": Найдено более двух устройств с мастерпинами. Остальные не обработаны."
                Exit For
            End If
            
            renameDevice.SetId sxtDevices(j)
            foundCount = j + 1
            
            e3App.PutInfo 0, "Строка " & rowNumber & ": --- Обработка устройства #" & foundCount & " ---"
            
            ' Переименование устройства
            newName = "-XT666"
            resultSet = renameDevice.SetName(newName)
            
            If resultSet = 0 Then
                e3App.PutInfo 0, "Строка " & rowNumber & ": Ошибка при переименовании устройства #" & foundCount
            Else
                e3App.PutInfo 0, "Строка " & rowNumber & ": Устройство #" & foundCount & " переименовано в " & newName
            End If
            
            ' Установка мастерпина
            If foundCount = 1 Then
                ' Используем значение из столбца I для первого устройства
                If Len(Trim(masterPin1Value)) > 0 Then
                    resultPin = renameDevice.SetMasterPinName(masterPin1Value)
                    If resultPin = 0 Then
                        e3App.PutInfo 0, "Строка " & rowNumber & ": Ошибка при установке мастерпина '" & masterPin1Value & "' для устройства #" & foundCount
                    Else
                        e3App.PutInfo 0, "Строка " & rowNumber & ": Мастерпин устройства #" & foundCount & " установлен в: " & masterPin1Value
                    End If
                Else
                    e3App.PutInfo 1, "Строка " & rowNumber & ": Значение мастерпина для первого устройства (столбец I) пустое. Мастерпин не изменен."
                End If
            ElseIf foundCount = 2 Then
                ' Используем значение из столбца J для второго устройства
                If Len(Trim(masterPin2Value)) > 0 Then
                    resultPin = renameDevice.SetMasterPinName(masterPin2Value)
                    If resultPin = 0 Then
                        e3App.PutInfo 0, "Строка " & rowNumber & ": Ошибка при установке мастерпина '" & masterPin2Value & "' для устройства #" & foundCount
                    Else
                        e3App.PutInfo 0, "Строка " & rowNumber & ": Мастерпин устройства #" & foundCount & " установлен в: " & masterPin2Value
                    End If
                Else
                    e3App.PutInfo 1, "Строка " & rowNumber & ": Значение мастерпина для второго устройства (столбец J) пустое. Мастерпин не изменен."
                End If
            End If
            
            ' Проверка результата
            finalMasterPin = renameDevice.GetMasterPinName()
            e3App.PutInfo 0, "Строка " & rowNumber & ": Итоговый мастерпин: '" & finalMasterPin & "'"
        Next
        
        If sxtCount = 0 Then
            e3App.PutInfo 0, "Строка " & rowNumber & ": Не найдено ни одного устройства -sXT1 с мастерпином."
        End If
        
    Else
        e3App.PutInfo 0, "Строка " & rowNumber & ": Ошибка: устройства в проекте не найдены."
    End If
    
    e3App.PutInfo 0, "Строка " & rowNumber & ": === ПЕРЕИМЕНОВАНИЕ ЗАВЕРШЕНО ==="
    
    ' Очистка
    Set renameDevice = Nothing
    On Error GoTo 0
End Sub

' --- Вспомогательные подпрограммы для очистки объектов ---
Sub CleanUpExcelObjects(ByRef objSheet, ByRef objWorkbook, ByRef objExcel)
    On Error Resume Next
    If Not objWorkbook Is Nothing Then objWorkbook.Close False
    If Not objExcel Is Nothing Then objExcel.Quit
    Set objSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0
End Sub

Sub CleanUpE3Objects(ByRef pin, ByRef device, ByRef job, ByRef sheet, ByRef e3App)
    On Error Resume Next
    Set pin = Nothing
    Set device = Nothing
    Set job = Nothing
    Set sheet = Nothing
    Set e3App = Nothing
    On Error GoTo 0
End Sub