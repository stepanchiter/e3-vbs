Option Explicit

' --- Глобальные переменные ---
' Объект приложения E3.series
Dim e3App
' Объект Job, представляющий текущий проект
Dim job
' Словарь для хранения ID найденных OOS символов, соответствующих критериям.
' Ключ: численное значение из имени OOS символа (например, "123" для "OOS123")
' Значение: ID символа в E3.series
Dim global_foundOOSIds

' Словарь для хранения соответствий компонента и диапазона E_Inom
' Ключ: Имя компонента (String)
' Значение: Массив Double(2) - [МинимальноеЗначение, МаксимальноеЗначение]
Dim componentMap

' --- Основная процедура скрипта ---
Sub Main()
    ' Инициализация объектов E3.series
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()

    ' Инициализация глобального словаря для хранения найденных OOS символов
    Set global_foundOOSIds = CreateObject("Scripting.Dictionary")
    
    ' Инициализация и заполнение словаря соответствия компонентов
    Set componentMap = CreateObject("Scripting.Dictionary")
    Call PopulateComponentMap() ' Вызываем процедуру для заполнения таблицы соответствий

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Поиск OOS символов и связанных с ними устройств ==="

    ' Шаг 1: Находим и фиксируем OOS символы по заданным критериям
    Call FindAndLogOOSSymbols()

    ' Шаг 2: Находим и выводим информацию о связанных устройствах (-QF и -KM)
    Call FindAndLogRelatedDevices()

    ' Шаг 3: Обновляем компонент QF на основе атрибута OOS символа
    Call UpdateQFComponentsBasedOnOOSAttribute()

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА ==="

    ' Очистка глобальных объектов для освобождения ресурсов
    Call CleanUpGlobalObjects()
End Sub

' --- Процедура для заполнения словаря соответствия компонентов ---
Sub PopulateComponentMap()
    ' Для каждого элемента словаря: Add "ИмяКомпонента", Array(МинимальноеЗначение, МаксимальноеЗначение)
    ' Используем CDbl() для явного преобразования чисел с плавающей точкой.
    ' Максимальное значение в диапазоне "от X до Y" здесь включительно Y-0.0001 для корректной работы If-ElseIf
    
    ' Автомат_3P_0.16-0.25 - от 0,16 до 0,2499;  
    componentMap.Add "Автомат_3P_0.16-0.25A", Array(CDbl(0.16), CDbl(0.2499))
    ' Автомат_3P_0.25-0.4 - от 0,25 до 0,3999;  
    componentMap.Add "Автомат_3P_0.25-0.4A", Array(CDbl(0.25), CDbl(0.3999))
    ' Автомат_3P_0.4-0.63 - от 0,40 до 0,6299;  
    componentMap.Add "Автомат_3P_0.4-0.63A", Array(CDbl(0.40), CDbl(0.6299))
    ' Автомат_3P_0.63-1.0 - от 0,63 до 0,9999;  
    componentMap.Add "Автомат_3P_0.63-1.0A", Array(CDbl(0.63), CDbl(0.9999))
    ' Автомат_3P_1.0-1.6 - от 1,00 до 1,5999;  
    componentMap.Add "Автомат_3P_1.0-1.6A", Array(CDbl(1.00), CDbl(1.5999))
    ' Автомат_3P_1.6-2.5 - от 1,60 до 2,4999;  
    componentMap.Add "Автомат_3P_1.6-2.5A", Array(CDbl(1.60), CDbl(2.4999))
    ' Автомат_3P_2.5-4.0 - от 2,50 до 3,9999;  
    componentMap.Add "Автомат_3P_2.5-4.0A", Array(CDbl(2.50), CDbl(3.9999))
    ' Автомат_3P_4.0-6.3 - от 4,00 до 6,2999;  
    componentMap.Add "Автомат_3P_4.0-6.3A", Array(CDbl(4.00), CDbl(6.2999))
    ' Автомат_3P_6.3-10.0 - от 6,30 до 9,9999;  
    componentMap.Add "Автомат_3P_6.3-10.0A", Array(CDbl(6.30), CDbl(9.9999))
    ' Автомат_3P_9-14 - от 9,00 до 13,9999;  
    componentMap.Add "Автомат_3P_9-14A", Array(CDbl(9.00), CDbl(13.9999))
    ' Автомат_3P_13-18 - от 13,00 до 17,9999;  
    componentMap.Add "Автомат_3P_13-18A", Array(CDbl(13.00), CDbl(17.9999))
    ' Автомат_3P_17-23 - от 17,00 до 22,9999;  
    componentMap.Add "Автомат_3P_17-23A", Array(CDbl(17.00), CDbl(22.9999))
    ' Автомат_3P_20-25 - от 20,00 до 24,9999;  
    componentMap.Add "Автомат_3P_20-25A", Array(CDbl(20.00), CDbl(24.9999))
    ' Автомат_3P_24-32 - от 24,00 до 31,9999;  
    componentMap.Add "Автомат_3P_24-32A", Array(CDbl(24.00), CDbl(31.9999))
    ' Автомат_3P_25-40 - от 25,00 до 39,9999;  
    componentMap.Add "Автомат_3P_25-40A", Array(CDbl(25.00), CDbl(39.9999))
    ' Автомат_3P_40-63 - от 40,00 до 62,9999;  
    componentMap.Add "Автомат_3Р_40-63А", Array(CDbl(40.00), CDbl(62.9999))
    ' Автомат_3P_56-80 - от 56,00 до 79,9999; 
    componentMap.Add "Автомат_3P_56-80A", Array(CDbl(56.00), CDbl(79.9999))
    
    e3App.PutInfo 0, "Загружено " & componentMap.Count & " соответствий компонентов."
End Sub

' --- Процедура для поиска и вывода информации об OOS символах ---
Sub FindAndLogOOSSymbols()
    Dim symbol          ' Объект Symbol для работы с отдельными символами
    Dim allSymbolIds()  ' Массив для хранения идентификаторов всех символов в проекте
    Dim allSymbolCount  ' Общее количество символов в проекте
    Dim i               ' Счетчик цикла для перебора символов

    Dim symbolName      ' Имя текущего символа
    Dim dProizv3Value   ' Значение атрибута "ОД D_Proizv3" текущего символа
    Dim OOSIndex        ' Числовой индекс из имени OOS символа (например, 123 для "OOS123")

    ' Создаем объект Symbol
    Set symbol = job.CreateSymbolObject()

    ' Получаем список всех символов в текущем проекте
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    ' Проверяем, есть ли символы в проекте
    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "В текущем проекте не найдено символов для анализа."
        Set symbol = Nothing ' Освобождаем объект Symbol перед выходом
        Exit Sub
    End If  
    
    e3App.PutInfo 0, "Найдено " & allSymbolCount & " символов в проекте. Ищем OOS символы с 'ОД D_Proizv3' = '3' или '4'..."

    Dim foundOOSCount : foundOOSCount = 0 ' Счетчик найденных OOS символов, соответствующих критериям

    ' Перебираем все символы в проекте
    For i = 1 To allSymbolCount
        ' Устанавливаем текущий символ по его ID для дальнейшей работы
        symbol.SetId(allSymbolIds(i))
        symbolName = symbol.GetName() ' Получаем имя символа

        ' Проверяем, начинается ли имя символа с "OOS" (без учета регистра)
        If LCase(Left(symbolName, 3)) = "OOS" Then
            ' Получаем значение атрибута "ОД D_Proizv3"
            ' Trim() удаляет лишние пробелы, CStr() преобразует в строку для надежного сравнения
            dProizv3Value = Trim(CStr(symbol.GetAttributeValue("ОД D_Proizv3")))

            ' Проверяем, соответствует ли значение атрибута нашим критериям ("3" или "4")
            If dProizv3Value = "3" Or dProizv3Value = "4" Or dProizv3Value = "8" Then
                foundOOSCount = foundOOSCount + 1 ' Увеличиваем счетчик
                
                ' Извлекаем числовой индекс из имени OOS символа (например, "123" из "OOS123")
                On Error Resume Next ' Включаем обработку ошибок для CLng
                OOSIndex = CLng(Mid(symbolName, 4)) ' Пытаемся преобразовать часть имени в число
                If Err.Number <> 0 Then
                    ' Если преобразование не удалось (например, "OOSABC"), используем исходную строку
                    OOSIndex = Mid(symbolName, 4)
                    e3App.PutInfo 0, "    ВНИМАНИЕ: Не удалось преобразовать индекс '" & Mid(symbolName, 4) & "' в число для OOS символа '" & symbolName & "'."
                    Err.Clear ' Очищаем ошибку
                End If
                On Error GoTo 0 ' Отключаем обработку ошибок

                ' Добавляем найденный символ в глобальный словарь
                ' Используем CStr(OOSIndex) для ключа, чтобы быть уверенными в типе данных ключа
                If Not global_foundOOSIds.Exists(CStr(OOSIndex)) Then
                    global_foundOOSIds.Add CStr(OOSIndex), allSymbolIds(i)
                    e3App.PutInfo 0, "  Найден и добавлен OOS символ: '" & symbolName & "'" & _
                                     " (ID: " & allSymbolIds(i) & ")" & _
                                     " | Атрибут 'ОД D_Proizv3': '" & dProizv3Value & "'"
                Else
                    e3App.PutInfo 0, "  ДУБЛИКАТ: OOS символ с индексом '" & CStr(OOSIndex) & "' уже найден. Обновляем ID на: " & allSymbolIds(i) & _
                                     " (Имя: '" & symbolName & "', D_Proizv3: '" & dProizv3Value & "')"
                    global_foundOOSIds.Item(CStr(OOSIndex)) = allSymbolIds(i) ' Обновляем ID, если такой индекс уже есть
                End If
            End If
        End If
    Next

    ' Выводим итоговое сообщение о результатах поиска OOS символов
    If foundOOSCount = 0 Then
        e3App.PutInfo 0, "Не найдено OOS символов со значением атрибута 'ОД D_Proizv3' равным '3' или '4'."
    Else
        e3App.PutInfo 0, "Всего найдено " & foundOOSCount & " OOS символов, соответствующих заданным критериям."
        e3App.PutInfo 0, "ID найденных OOS символов сохранены в глобальном словаре 'global_foundOOSIds'."
    End If

    Set symbol = Nothing ' Освобождаем объект Symbol
End Sub

' --- Процедура для поиска и вывода информации о связанных устройствах (-QF и -KM) ---
Sub FindAndLogRelatedDevices()
    Dim device          ' Объект Device для работы с устройствами
    Dim OOSIndex_str    ' Строковое представление числового индекса OOS символа
    Dim targetDeviceName    ' Имя устройства, которое мы ищем (например, "-QF123" или "-KM123")
    Dim allDeviceIds()      ' Массив для хранения идентификаторов всех устройств в проекте
    Dim allDeviceCount      ' Общее количество устройств в проекта
    Dim i                   ' Счетчик цикла для перебора устройств
    Dim currentDeviceName   ' Имя текущего устройства
    Dim componentName       ' Имя компонента текущего устройства

    ' Проверяем, были ли найдены OOS символы на предыдущем шаге
    If global_foundOOSIds.Count = 0 Then
        e3App.PutInfo 0, "Нет зафиксированных OOS символов (D_Proizv3=3 или 4) для поиска связанных устройств."
        Exit Sub
    End If

    Set device = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== Начало поиска связанных устройств -QF и -KM для найденных OOS символов ==="
    
    ' Получаем список всех устройств в проекте один раз для эффективности
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' Перебираем каждый зафиксированный OOS символ
    For Each OOSIndex_str In global_foundOOSIds.Keys
        e3App.PutInfo 0, "  Поиск связанных устройств для OOS" & OOSIndex_str & ":"
        Dim foundRelatedDeviceForCurrentOOS : foundRelatedDeviceForCurrentOOS = False

        ' --- Поиск -QF устройств ---
        targetDeviceName = "-QF" & OOSIndex_str
        Dim qfFoundCount : qfFoundCount = 0 ' Счетчик найденных валидных -QF устройств
        
        For i = 1 To allDeviceCount ' Перебираем ВСЕ устройства, чтобы найти все совпадения
            device.SetId(allDeviceIds(i))
            currentDeviceName = device.GetName()
            componentName = device.GetComponentName()

            If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                ' Для -QF устройств, также проверяем, содержит ли компонент "Автомат"
                If InStr(1, LCase(componentName), "автомат") > 0 Then
                    qfFoundCount = qfFoundCount + 1
                    e3App.PutInfo 0, "    Найдено -QF устройство: '" & currentDeviceName & "'" & _
                                     " (ID: " & allDeviceIds(i) & ")" & _
                                     " | Компонент: '" & componentName & "'"
                    foundRelatedDeviceForCurrentOOS = True
                Else
                    e3App.PutInfo 0, "    Найденo -QF устройство: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), но его компонент ('" & componentName & "') не содержит 'Автомат'. Это устройство пропущено."
                End If
            End If
        Next
        
        If qfFoundCount = 0 Then
            e3App.PutInfo 0, "    -QF" & OOSIndex_str & " (с компонентом 'Автомат') не найдено ни одного устройства среди всех устройств проекта."
        Else
            e3App.PutInfo 0, "    Всего найдено " & qfFoundCount & " -QF устройств с компонентом 'Автомат' для OOS" & OOSIndex_str & "."
        End If

        ' --- Поиск -KM устройств ---
        targetDeviceName = "-KM" & OOSIndex_str
        Dim kmFoundCount : kmFoundCount = 0 ' Счетчик найденных -KM устройств

        For i = 1 To allDeviceCount ' Перебираем ВСЕ устройства, чтобы найти все совпадения
            device.SetId(allDeviceIds(i))
            currentDeviceName = device.GetName()
            componentName = device.GetComponentName()

            If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                kmFoundCount = kmFoundCount + 1
                e3App.PutInfo 0, "    Найдено -KM устройство: '" & currentDeviceName & "'" & _
                                 " (ID: " & allDeviceIds(i) & ")" & _
                                 " | Компонент: '" & componentName & "'"
                foundRelatedDeviceForCurrentOOS = True
            End If
        Next
        
        If kmFoundCount = 0 Then
            e3App.PutInfo 0, "    -KM" & OOSIndex_str & " не найдено ни одного устройства среди всех устройств проекта."
        Else
            e3App.PutInfo 0, "    Всего найдено " & kmFoundCount & " -KM устройств для OOS" & OOSIndex_str & "."
        End If

        If Not foundRelatedDeviceForCurrentOOS Then
            e3App.PutInfo 0, "  Для OOS" & OOSIndex_str & " не найдено ни одного соответствующего -QF (с компонентом 'Автомат') или -KM устройства."
        End If
    Next

    e3App.PutInfo 0, "=== Завершение поиска связанных устройств ==="

    Set device = Nothing ' Освобождаем объект Device
End Sub

' --- Процедура для обновления компонента QF на основе атрибута OOS символа ---
Sub UpdateQFComponentsBasedOnOOSAttribute()
    Dim symbolObj       ' Объект Symbol для чтения атрибутов OOS
    Dim deviceObj       ' Объект Device для обновления компонентов QF
    Dim OOSIndex_str    ' Строковое представление числового индекса OOS символа
    Dim OOSSymbolId     ' ID OOS символа
    Dim eInomValue_str  ' Строковое значение атрибута "ОД E_Inom" (исходное)
    Dim eInomValue_num  ' Числовое значение атрибута "ОД E_Inom"
    Dim isEInomValueValid ' Флаг для проверки успешности преобразования
    
    Dim targetDeviceName_QF ' Ожидаемое имя QF устройства
    Dim allDeviceIds()      ' Массив ID всех устройств
    Dim allDeviceCount      ' Количество всех устройств
    Dim i                   ' Счетчик цикла

    ' НОВЫЕ ПЕРЕМЕННЫЕ ДЛЯ ПОИСКА КОМПОНЕНТА
    Dim componentName_to_set ' Имя компонента, который нужно установить
    Dim rangeValues          ' Массив с мин/макс значениями для текущего компонента из словаря
    Dim foundMatchingComponent ' Флаг, указывающий, найден ли подходящий компонент
    Dim componentName_key    ' Переменная для перебора ключей словаря

    ' Константа для версии компонента
    Const COMPONENT_VERSION = "1" ' Версия компонента

    ' Проверяем, были ли найдены OOS символы на предыдущем шаге
    If global_foundOOSIds.Count = 0 Then
        e3App.PutInfo 0, "COMM: Нет зафиксированных OOS символов (D_Proizv3=3 или 4) для обновления компонентов QF."
        Exit Sub
    End If

    Set symbolObj = job.CreateSymbolObject()
    Set deviceObj = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== Начало обновления компонентов -QF на основе атрибута 'ОД E_Inom' OOS символов ==="

    ' Проверяем, что словарь соответствия компонентов не пуст
    If componentMap.Count = 0 Then
        e3App.PutInfo 0, "ОШИБКА: Словарь соответствия компонентов пуст. Невозможно обновить компоненты."
        Exit Sub
    End If

    ' Получаем список всех устройств в проекте один раз для эффективности
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' Перебираем каждый зафиксированный OOS символ
    For Each OOSIndex_str In global_foundOOSIds.Keys
        OOSSymbolId = global_foundOOSIds.Item(OOSIndex_str)
        
        ' Устанавливаем OOS символ для чтения атрибута
        symbolObj.SetId(OOSSymbolId)
        
        ' Применяем рабочий подход для получения и преобразования значения E_Inom
        eInomValue_str = CStr(symbolObj.GetAttributeValue("ОД E_Inom")) 

        e3App.PutInfo 0, "  Обработка OOS" & OOSIndex_str & " (ID: " & OOSSymbolId & ")"
        e3App.PutInfo 0, "    Исходный атрибут 'ОД E_Inom': '" & eInomValue_str & "'"
        e3App.PutInfo 0, "    Атрибут 'ОД E_Inom' после Trim(): '" & Trim(eInomValue_str) & "'"

        isEInomValueValid = False ' Изначально считаем невалидным
        
        Dim trimmedEInomValue_str : trimmedEInomValue_str = Trim(eInomValue_str)

        If IsNumeric(trimmedEInomValue_str) And Len(trimmedEInomValue_str) > 0 Then
            On Error Resume Next ' Включаем обработку ошибок для CDbl
            eInomValue_num = CDbl(trimmedEInomValue_str)
            If Err.Number = 0 Then
                isEInomValueValid = True ' Преобразование успешно
                e3App.PutInfo 0, "    УСПЕШНО: Преобразованное числовое значение 'ОД E_Inom': " & eInomValue_num
            Else
                e3App.PutInfo 0, "    ОШИБКА: CDbl не удалось преобразовать строку '" & trimmedEInomValue_str & "' в число (Err: " & Err.Description & ")"
                Err.Clear ' Очищаем ошибку
            End If
            On Error GoTo 0 ' Отключаем обработку ошибок
        Else
            e3App.PutInfo 0, "    ВНИМАНИЕ: Атрибут 'ОД E_Inom' ('" & eInomValue_str & "') пуст или не является числом. Пропускаем обновление."
        End If

        ' Только если преобразование прошло успешно, ищем подходящий компонент
        If isEInomValueValid Then
            foundMatchingComponent = False
            componentName_to_set = "" ' Сбрасываем для каждого OOS символа

            ' Перебираем словарь в поисках подходящего диапазона
            For Each componentName_key In componentMap.Keys
                rangeValues = componentMap.Item(componentName_key) ' Получаем массив [min, max]
                
                If eInomValue_num >= rangeValues(0) And eInomValue_num <= rangeValues(1) Then
                    componentName_to_set = componentName_key ' Нашли подходящее имя компонента
                    foundMatchingComponent = True
                    e3App.PutInfo 0, "    Найдено подходящее имя компонента: '" & componentName_to_set & "' для значения " & eInomValue_num
                    Exit For ' Выходим из цикла, так как нашли первое совпадение
                End If
            Next

            If foundMatchingComponent Then
                e3App.PutInfo 0, "    Поиск связанных -QF устройств для обновления компонента на: '" & componentName_to_set & "'..."
                
                targetDeviceName_QF = "-QF" & OOSIndex_str
                Dim qfUpdatedCount : qfUpdatedCount = 0 

                For i = 1 To allDeviceCount
                    deviceObj.SetId(allDeviceIds(i))
                    Dim currentDeviceName : currentDeviceName = deviceObj.GetName()
                    Dim currentComponentName : currentComponentName = deviceObj.GetComponentName()

                    If UCase(currentDeviceName) = UCase(targetDeviceName_QF) Then
                        ' Дополнительная проверка, что компонент QF содержит "Автомат"
                        If InStr(1, LCase(currentComponentName), "автомат") > 0 Then
                            e3App.PutInfo 0, "      Найдено -QF устройство для обновления: '" & currentDeviceName & "'" & _
                                             " (ID: " & allDeviceIds(i) & ", Текущий компонент: '" & currentComponentName & "')"
                            
                            On Error Resume Next ' Включаем обработку ошибок для SetComponentName
                            deviceObj.SetComponentName componentName_to_set, COMPONENT_VERSION
                            If Err.Number = 0 Then
                                qfUpdatedCount = qfUpdatedCount + 1
                                e3App.PutInfo 0, "        УСПЕШНО: Компонент обновлен на: '" & componentName_to_set & "' (Версия: '" & COMPONENT_VERSION & "')."
                            Else
                                e3App.PutInfo 0, "        ОШИБКА при обновлении компонента для '" & currentDeviceName & "': " & Err.Description
                                Err.Clear ' Очищаем ошибку
                            End If
                            On Error GoTo 0 ' Отключаем обработку ошибок
                        Else
                            e3App.PutInfo 0, "      Найденo -QF устройство: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), но его компонент ('" & currentComponentName & "') не содержит 'Автомат'. Пропущено обновление."
                        End If
                    End If
                Next
                
                If qfUpdatedCount = 0 Then
                    e3App.PutInfo 0, "    Для OOS" & OOSIndex_str & " не найдено ни одного -QF устройства с компонентом 'Автомат' для обновления."
                Else
                    e3App.PutInfo 0, "    Всего обновлено " & qfUpdatedCount & " -QF устройств для OOS" & OOSIndex_str & "."
                End If
            Else
                e3App.PutInfo 0, "    ВНИМАНИЕ: Для значения 'ОД E_Inom' (" & eInomValue_num & ") не найдено подходящего компонента в таблице соответствия. Обновление пропущено."
            End If
        End If 
    Next ' Продолжаем к следующему OOS символу

    e3App.PutInfo 0, "=== Завершение обновления компонентов -QF ==="

    Set symbolObj = Nothing ' Освобождаем объект Symbol
    Set deviceObj = Nothing ' Освобождаем объект Device
End Sub


' --- Вспомогательная процедура для очистки глобальных объектов ---
Sub CleanUpGlobalObjects()
    ' Проверяем, что объекты существуют, прежде чем их освобождать
    If Not job Is Nothing Then
        Set job = Nothing
    End If
    If Not e3App Is Nothing Then
        Set e3App = Nothing
    End If
    If Not global_foundOOSIds Is Nothing Then
        Set global_foundOOSIds = Nothing
    End If
    ' Освобождаем объект componentMap
    If Not componentMap Is Nothing Then
        Set componentMap = Nothing
    End If
End Sub

' --- Точка входа в скрипт: запускаем основную процедуру ---
Call Main()