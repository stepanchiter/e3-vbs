'*******************************************************************************
' Название скрипта: E3_UZ_ComponentUpdater
' Автор: E3.series VBScript Assistant
' Дата: 01.07.2025
' Описание: Скрипт для автоматического обновления имен компонентов для устройств -QF и -KM
'           на основе артикулов, извлеченных из OOO символов, и новой таблицы соответствия.
'           Не проверяет имя компонента на "автомат" для KM, только для QF.
'*******************************************************************************
Option Explicit

' --- Глобальные переменные ---
' Объект приложения E3.series
Dim e3App
' Объект Job, представляющий текущий проект
Dim job
' Словарь для хранения ID найденных OOO символов, соответствующих критериям.
' Ключ: численное значение из имени OOO символа (например, "123" для "OOO123")
' Значение: ID символа в E3.series
Dim global_foundOOOIds

' Словарь для хранения соответствий компонента и диапазона E_Inom для Автоматов (QF)
' Ключ: Имя компонента (String)
' Значение: Массив Double(2) - [МинимальноеЗначение, МаксимальноеЗначение]
Dim qfComponentMap

' НОВЫЙ СЛОВАРЬ для хранения соответствий компонента и диапазона E_Inom для Контакторов (KM)
Dim kmComponentMap

' --- Основная процедура скрипта ---
Sub Main()
    ' Инициализация объектов E3.series
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()

    ' Инициализация глобального словаря для хранения найденных OOO символов
    Set global_foundOOOIds = CreateObject("Scripting.Dictionary")
    
    ' Инициализация и заполнение словаря соответствия компонентов QF
    Set qfComponentMap = CreateObject("Scripting.Dictionary")
    Call PopulateQFComponentMap() ' Вызываем процедуру для заполнения таблицы соответствий QF

    ' НОВОЕ: Инициализация и заполнение словаря соответствия компонентов KM
    Set kmComponentMap = CreateObject("Scripting.Dictionary")
    Call PopulateKMComponentMap() ' Вызываем процедуру для заполнения таблицы соответствий KM

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Поиск OOO символов и связанных с ними устройств ==="

    ' Шаг 1: Находим и фиксируем OOO символы по заданным критериям
    Call FindAndLogOOOSymbols()

    ' Шаг 2: Находим и выводим информацию о связанных устройствах (-QF и -KM)
    Call FindAndLogRelatedDevices()

    ' Шаг 3: Обновляем компонент QF на основе атрибута OOO символа
    Call UpdateQFComponentsBasedOnOOOAttribute()

    ' НОВОЕ: Шаг 4: Обновляем компонент KM на основе атрибута OOO символа
    Call UpdateKMComponentsBasedOnOOOAttribute()

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА ==="

    ' Очистка глобальных объектов для освобождения ресурсов
    Call CleanUpGlobalObjects()
End Sub

' --- Процедура для заполнения словаря соответствия компонентов QF (Автоматов) ---
Sub PopulateQFComponentMap()
    ' Для каждого элемента словаря: Add "ИмяКомпонента", Array(МинимальноеЗначение, МаксимальноеЗначение)
    ' Используем CDbl() для явного преобразования чисел с плавающей точкой.
    ' Максимальное значение в диапазоне "от X до Y" здесь включительно Y-0.0001 для корректной работы If-ElseIf
    
    qfComponentMap.Add "Автомат_3P_0.16-0.25A", Array(CDbl(0.16), CDbl(0.2499))
    qfComponentMap.Add "Автомат_3P_0.25-0.4A", Array(CDbl(0.25), CDbl(0.3999))
    qfComponentMap.Add "Автомат_3P_0.4-0.63A", Array(CDbl(0.40), CDbl(0.6299))
    qfComponentMap.Add "Автомат_3P_0.63-1.0A", Array(CDbl(0.63), CDbl(0.9999))
    qfComponentMap.Add "Автомат_3P_1.0-1.6A", Array(CDbl(1.00), CDbl(1.5999))
    qfComponentMap.Add "Автомат_3P_1.6-2.5A", Array(CDbl(1.60), CDbl(2.4999))
    qfComponentMap.Add "Автомат_3P_2.5-4.0A", Array(CDbl(2.50), CDbl(3.9999))
    qfComponentMap.Add "Автомат_3P_4.0-6.3A", Array(CDbl(4.00), CDbl(6.2999))
    qfComponentMap.Add "Автомат_3P_6.3-10.0A", Array(CDbl(6.30), CDbl(9.9999))
    qfComponentMap.Add "Автомат_3P_9-14A", Array(CDbl(9.00), CDbl(13.9999))
    qfComponentMap.Add "Автомат_3P_13-18A", Array(CDbl(13.00), CDbl(17.9999))
    qfComponentMap.Add "Автомат_3P_17-23A", Array(CDbl(17.00), CDbl(22.9999))
    qfComponentMap.Add "Автомат_3P_20-25A", Array(CDbl(20.00), CDbl(24.9999))
    qfComponentMap.Add "Автомат_3P_24-32A", Array(CDbl(24.00), CDbl(31.9999))
    qfComponentMap.Add "Автомат_3P_25-40A", Array(CDbl(25.00), CDbl(39.9999))
    qfComponentMap.Add "Автомат_3Р_40-63А", Array(CDbl(40.00), CDbl(62.9999))
    qfComponentMap.Add "Автомат_3P_56-80A", Array(CDbl(56.00), CDbl(79.9999))
    
    e3App.PutInfo 0, "Загружено " & qfComponentMap.Count & " соответствий компонентов QF (Автоматов)."
End Sub

' НОВАЯ ПРОЦЕДУРА: для заполнения словаря соответствия компонентов KM (Контакторов) ---
Sub PopulateKMComponentMap()
    kmComponentMap.Add "Контактор_КМ102_22001DEK", Array(CDbl(0.00), CDbl(9.00))
    kmComponentMap.Add "Контактор_КМ102_22002DEK", Array(CDbl(9.01), CDbl(12.00))
    kmComponentMap.Add "Контактор_КМ102_22003DEK", Array(CDbl(12.01), CDbl(18.00))
    kmComponentMap.Add "Контактор_КМ102_22004DEK", Array(CDbl(18.01), CDbl(25.00))
    kmComponentMap.Add "Контактор_КМ102_22005DEK", Array(CDbl(25.01), CDbl(32.00))
    kmComponentMap.Add "Контактор_КМ102_22006DEK", Array(CDbl(32.01), CDbl(40.00))
    kmComponentMap.Add "Контактор_КМ102_22007DEK", Array(CDbl(40.01), CDbl(50.00))
    kmComponentMap.Add "Контактор_КМ102_22008DEK", Array(CDbl(50.01), CDbl(65.00))
    kmComponentMap.Add "Контактор_КМ102_22009DEK", Array(CDbl(65.01), CDbl(80.00))
    kmComponentMap.Add "Контактор_КМ102_22010DEK", Array(CDbl(80.01), CDbl(95.00))
    kmComponentMap.Add "Контактор_КМ103_22150DEK", Array(CDbl(95.01), CDbl(115.00))
    kmComponentMap.Add "Контактор_КМ103_22152DEK", Array(CDbl(115.01), CDbl(150.00))
    kmComponentMap.Add "Контактор_КМ103_22154DEK", Array(CDbl(150.01), CDbl(185.00))
    kmComponentMap.Add "Контактор_КМ103_22156DEK", Array(CDbl(185.01), CDbl(225.00))
    kmComponentMap.Add "Контактор_КМ103_22158DEK", Array(CDbl(225.01), CDbl(265.00))
    kmComponentMap.Add "Контактор_КМ103_22160DEK", Array(CDbl(265.01), CDbl(330.00))
    kmComponentMap.Add "Контактор_КМ103_22162DEK", Array(CDbl(330.01), CDbl(400.00))

    e3App.PutInfo 0, "Загружено " & kmComponentMap.Count & " соответствий компонентов KM (Контакторов)."
End Sub

' --- Процедура для поиска и вывода информации об OOO символах ---
Sub FindAndLogOOOSymbols()
    Dim symbol            ' Объект Symbol для работы с отдельными символами
    Dim allSymbolIds()    ' Массив для хранения идентификаторов всех символов в проекте
    Dim allSymbolCount    ' Общее количество символов в проекте
    Dim i                 ' Счетчик цикла для перебора символов

    Dim symbolName        ' Имя текущего символа
    Dim dProizv3Value     ' Значение атрибута "ОД D_Proizv3" текущего символа
    Dim oooIndex          ' Числовой индекс из имени OOO символа (например, 123 для "OOO123")

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
    
    e3App.PutInfo 0, "Найдено " & allSymbolCount & " символов в проекте. Ищем OOO символы с 'ОД D_Proizv3' = '3', '4' или '8'..."

    Dim foundOOOCount : foundOOOCount = 0 ' Счетчик найденных OOO символов, соответствующих критериям

    ' Перебираем все символы в проекте
    For i = 1 To allSymbolCount
        ' Устанавливаем текущий символ по его ID для дальнейшей работы
        symbol.SetId(allSymbolIds(i))
        symbolName = symbol.GetName() ' Получаем имя символа

        ' Проверяем, начинается ли имя символа с "OOO" (без учета регистра)
        If LCase(Left(symbolName, 3)) = "ooo" Then
            ' Получаем значение атрибута "ОД D_Proizv3"
            ' Trim() удаляет лишние пробелы, CStr() преобразует в строку для надежного сравнения
            dProizv3Value = Trim(CStr(symbol.GetAttributeValue("ОД D_Proizv3")))

            ' Проверяем, соответствует ли значение атрибута нашим критериям ("3" или "4" или "8")
            If dProizv3Value = "3" Or dProizv3Value = "4" Or dProizv3Value = "8" Then
                foundOOOCount = foundOOOCount + 1 ' Увеличиваем счетчик
                
                ' Извлекаем числовой индекс из имени OOO символа (например, "123" из "OOO123")
                On Error Resume Next ' Включаем обработку ошибок для CLng
                oooIndex = CLng(Mid(symbolName, 4)) ' Пытаемся преобразовать часть имени в число
                If Err.Number <> 0 Then
                    ' Если преобразование не удалось (например, "OOOABC"), используем исходную строку
                    oooIndex = Mid(symbolName, 4)
                    e3App.PutInfo 0, "    ВНИМАНИЕ: Не удалось преобразовать индекс '" & Mid(symbolName, 4) & "' в число для OOO символа '" & symbolName & "'."
                    Err.Clear ' Очищаем ошибку
                End If
                On Error GoTo 0 ' Отключаем обработку ошибок

                ' Добавляем найденный символ в глобальный словарь
                ' Используем CStr(oooIndex) для ключа, чтобы быть уверенными в типе данных ключа
                If Not global_foundOOOIds.Exists(CStr(oooIndex)) Then
                    global_foundOOOIds.Add CStr(oooIndex), allSymbolIds(i)
                    e3App.PutInfo 0, "  Найден и добавлен OOO символ: '" & symbolName & "'" & _
                                     " (ID: " & allSymbolIds(i) & ")" & _
                                     " | Атрибут 'ОД D_Proizv3': '" & dProizv3Value & "'"
                Else
                    e3App.PutInfo 0, "  ДУБЛИКАТ: OOO символ с индексом '" & CStr(oooIndex) & "' уже найден. Обновляем ID на: " & allSymbolIds(i) & _
                                     " (Имя: '" & symbolName & "', D_Proizv3: '" & dProizv3Value & "')"
                    global_foundOOOIds.Item(CStr(oooIndex)) = allSymbolIds(i) ' Обновляем ID, если такой индекс уже есть
                End If
            End If
        End If
    Next

    ' Выводим итоговое сообщение о результатах поиска OOO символов
    If foundOOOCount = 0 Then
        e3App.PutInfo 0, "Не найдено OOO символов со значением атрибута 'ОД D_Proizv3' равным '3', '4' или '8'."
    Else
        e3App.PutInfo 0, "Всего найдено " & foundOOOCount & " OOO символов, соответствующих заданным критериям."
        e3App.PutInfo 0, "ID найденных OOO символов сохранены в глобальном словаре 'global_foundOOOIds'."
    End If

    Set symbol = Nothing ' Освобождаем объект Symbol
End Sub

' --- Процедура для поиска и вывода информации о связанных устройствах (-QF и -KM) ---
Sub FindAndLogRelatedDevices()
    Dim device            ' Объект Device для работы с устройствами
    Dim oooIndex_str      ' Строковое представление числового индекса OOO символа
    Dim targetDeviceName    ' Имя устройства, которое мы ищем (например, "-QF123" или "-KM123")
    Dim allDeviceIds()      ' Массив для хранения идентификаторов всех устройств в проекте
    Dim allDeviceCount      ' Общее количество устройств в проекта
    Dim i                   ' Счетчик цикла для перебора устройств
    Dim currentDeviceName   ' Имя текущего устройства
    Dim componentName       ' Имя компонента текущего устройства

    ' Проверяем, были ли найдены OOO символы на предыдущем шаге
    If global_foundOOOIds.Count = 0 Then
        e3App.PutInfo 0, "Нет зафиксированных OOO символов (D_Proizv3=3, 4 или 8) для поиска связанных устройств."
        Exit Sub
    End If

    Set device = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== Начало поиска связанных устройств -QF и -KM для найденных OOO символов ==="
    
    ' Получаем список всех устройств в проекте один раз для эффективности
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' Перебираем каждый зафиксированный OOO символ
    For Each oooIndex_str In global_foundOOOIds.Keys
        e3App.PutInfo 0, "  Поиск связанных устройств для OOO" & oooIndex_str & ":"
        Dim foundRelatedDeviceForCurrentOOO : foundRelatedDeviceForCurrentOOO = False

        ' --- Поиск -QF устройств ---
        targetDeviceName = "-QF" & oooIndex_str
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
                    foundRelatedDeviceForCurrentOOO = True
                Else
                    e3App.PutInfo 0, "    Найденo -QF устройство: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), но его компонент ('" & componentName & "') не содержит 'Автомат'. Это устройство пропущено."
                End If
            End If
        Next
        
        If qfFoundCount = 0 Then
            e3App.PutInfo 0, "    -QF" & oooIndex_str & " (с компонентом 'Автомат') не найдено ни одного устройства среди всех устройств проекта."
        Else
            e3App.PutInfo 0, "    Всего найдено " & qfFoundCount & " -QF устройств с компонентом 'Автомат' для OOO" & oooIndex_str & "."
        End If

        ' --- Поиск -KM устройств ---
        targetDeviceName = "-KM" & oooIndex_str
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
                foundRelatedDeviceForCurrentOOO = True
            End If
        Next
        
        If kmFoundCount = 0 Then
            e3App.PutInfo 0, "    -KM" & oooIndex_str & " не найдено ни одного устройства среди всех устройств проекта."
        Else
            e3App.PutInfo 0, "    Всего найдено " & kmFoundCount & " -KM устройств для OOO" & oooIndex_str & "."
        End If

        If Not foundRelatedDeviceForCurrentOOO Then
            e3App.PutInfo 0, "  Для OOO" & oooIndex_str & " не найдено ни одного соответствующего -QF (с компонентом 'Автомат') или -KM устройства."
        End If
    Next

    e3App.PutInfo 0, "=== Завершение поиска связанных устройств ==="

    Set device = Nothing ' Освобождаем объект Device
End Sub

' --- Процедура для обновления компонента QF на основе атрибута OOO символа ---
Sub UpdateQFComponentsBasedOnOOOAttribute()
    Dim symbolObj       ' Объект Symbol для чтения атрибутов OOO
    Dim deviceObj       ' Объект Device для обновления компонентов QF
    Dim oooIndex_str    ' Строковое представление числового индекса OOO символа
    Dim oooSymbolId     ' ID OOO символа
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

    ' Проверяем, были ли найдены OOO символы на предыдущем шаге
    If global_foundOOOIds.Count = 0 Then
        e3App.PutInfo 0, "COMM: Нет зафиксированных OOO символов (D_Proizv3=3, 4 или 8) для обновления компонентов QF."
        Exit Sub
    End If

    Set symbolObj = job.CreateSymbolObject()
    Set deviceObj = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== Начало обновления компонентов -QF на основе атрибута 'ОД E_Inom' OOO символов ==="

    ' Проверяем, что словарь соответствия компонентов QF не пуст
    If qfComponentMap.Count = 0 Then
        e3App.PutInfo 0, "ОШИБКА: Словарь соответствия компонентов QF пуст. Невозможно обновить компоненты QF."
        Exit Sub
    End If

    ' Получаем список всех устройств в проекте один раз для эффективности
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' Перебираем каждый зафиксированный OOO символ
    For Each oooIndex_str In global_foundOOOIds.Keys
        oooSymbolId = global_foundOOOIds.Item(oooIndex_str)
        
        ' Устанавливаем OOO символ для чтения атрибута
        symbolObj.SetId(oooSymbolId)
        
        ' Применяем рабочий подход для получения и преобразования значения E_Inom
        eInomValue_str = CStr(symbolObj.GetAttributeValue("ОД E_Inom"))    

        e3App.PutInfo 0, "  Обработка OOO" & oooIndex_str & " (ID: " & oooSymbolId & ") для QF"
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
            e3App.PutInfo 0, "    ВНИМАНИЕ: Атрибут 'ОД E_Inom' ('" & eInomValue_str & "') пуст или не является числом. Пропускаем обновление QF."
        End If

        ' Только если преобразование прошло успешно, ищем подходящий компонент
        If isEInomValueValid Then
            foundMatchingComponent = False
            componentName_to_set = "" ' Сбрасываем для каждого OOO символа

            ' Перебираем словарь QF в поисках подходящего диапазона
            For Each componentName_key In qfComponentMap.Keys
                rangeValues = qfComponentMap.Item(componentName_key) ' Получаем массив [min, max]
                
                If eInomValue_num >= rangeValues(0) And eInomValue_num <= rangeValues(1) Then
                    componentName_to_set = componentName_key ' Нашли подходящее имя компонента
                    foundMatchingComponent = True
                    e3App.PutInfo 0, "    Найдено подходящее имя компонента QF: '" & componentName_to_set & "' для значения " & eInomValue_num
                    Exit For ' Выходим из цикла, так как нашли первое совпадение
                End If
            Next

            If foundMatchingComponent Then
                e3App.PutInfo 0, "    Поиск связанных -QF устройств для обновления компонента на: '" & componentName_to_set & "'..."
                
                targetDeviceName_QF = "-QF" & oooIndex_str
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
                                e3App.PutInfo 0, "        УСПЕШНО: Компонент QF обновлен на: '" & componentName_to_set & "' (Версия: '" & COMPONENT_VERSION & "')."
                            Else
                                e3App.PutInfo 0, "        ОШИБКА при обновлении компонента QF для '" & currentDeviceName & "': " & Err.Description
                                Err.Clear ' Очищаем ошибку
                            End If
                            On Error GoTo 0 ' Отключаем обработку ошибок
                        Else
                            e3App.PutInfo 0, "      Найденo -QF устройство: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), но его компонент ('" & currentComponentName & "') не содержит 'Автомат'. Пропущено обновление."
                        End If
                    End If
                Next
                
                If qfUpdatedCount = 0 Then
                    e3App.PutInfo 0, "    Для OOO" & oooIndex_str & " не найдено ни одного -QF устройства с компонентом 'Автомат' для обновления."
                Else
                    e3App.PutInfo 0, "    Всего обновлено " & qfUpdatedCount & " -QF устройств для OOO" & oooIndex_str & "."
                End If
            Else
                e3App.PutInfo 0, "    ВНИМАНИЕ: Для значения 'ОД E_Inom' (" & eInomValue_num & ") не найдено подходящего компонента в таблице соответствия QF. Обновление пропущено."
            End If
        End If    
    Next ' Продолжаем к следующему OOO символу

    e3App.PutInfo 0, "=== Завершение обновления компонентов -QF ==="

    Set symbolObj = Nothing ' Освобождаем объект Symbol
    Set deviceObj = Nothing ' Освобождаем объект Device
End Sub

' НОВАЯ ПРОЦЕДУРА: для обновления компонента KM на основе атрибута OOO символа ---
Sub UpdateKMComponentsBasedOnOOOAttribute()
    Dim symbolObj       ' Объект Symbol для чтения атрибутов OOO
    Dim deviceObj       ' Объект Device для обновления компонентов KM
    Dim oooIndex_str    ' Строковое представление числового индекса OOO символа
    Dim oooSymbolId     ' ID OOO символа
    Dim eInomValue_str  ' Строковое значение атрибута "ОД E_Inom" (исходное)
    Dim eInomValue_num  ' Числовое значение атрибута "ОД E_Inom"
    Dim isEInomValueValid ' Флаг для проверки успешности преобразования
    
    Dim targetDeviceName_KM ' Ожидаемое имя KM устройства
    Dim allDeviceIds()      ' Массив ID всех устройств
    Dim allDeviceCount      ' Количество всех устройств
    Dim i                   ' Счетчик цикла

    Dim componentName_to_set_km ' Имя компонента KM, который нужно установить
    Dim rangeValues_km          ' Массив с мин/макс значениями для текущего KM компонента из словаря
    Dim foundMatchingComponent_km ' Флаг, указывающий, найден ли подходящий KM компонент
    Dim componentName_key_km    ' Переменная для перебора ключей словаря KM

    Const COMPONENT_VERSION = "1" ' Версия компонента

    If global_foundOOOIds.Count = 0 Then
        e3App.PutInfo 0, "COMM: Нет зафиксированных OOO символов (D_Proizv3=3, 4 или 8) для обновления компонентов KM."
        Exit Sub
    End If

    Set symbolObj = job.CreateSymbolObject()
    Set deviceObj = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== Начало обновления компонентов -KM на основе атрибута 'ОД E_Inom' OOO символов ==="

    If kmComponentMap.Count = 0 Then
        e3App.PutInfo 0, "ОШИБКА: Словарь соответствия компонентов KM пуст. Невозможно обновить компоненты KM."
        Exit Sub
    End If

    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    For Each oooIndex_str In global_foundOOOIds.Keys
        oooSymbolId = global_foundOOOIds.Item(oooIndex_str)
        
        symbolObj.SetId(oooSymbolId)
        eInomValue_str = CStr(symbolObj.GetAttributeValue("ОД E_Inom"))    

        e3App.PutInfo 0, "  Обработка OOO" & oooIndex_str & " (ID: " & oooSymbolId & ") для KM"
        e3App.PutInfo 0, "    Исходный атрибут 'ОД E_Inom': '" & eInomValue_str & "'"

        isEInomValueValid = False
        Dim trimmedEInomValue_str_km : trimmedEInomValue_str_km = Trim(eInomValue_str)

        If IsNumeric(trimmedEInomValue_str_km) And Len(trimmedEInomValue_str_km) > 0 Then
            On Error Resume Next
            eInomValue_num = CDbl(trimmedEInomValue_str_km)
            If Err.Number = 0 Then
                isEInomValueValid = True
                e3App.PutInfo 0, "    УСПЕШНО: Преобразованное числовое значение 'ОД E_Inom': " & eInomValue_num
            Else
                e3App.PutInfo 0, "    ОШИБКА: CDbl не удалось преобразовать строку '" & trimmedEInomValue_str_km & "' в число (Err: " & Err.Description & ")"
                Err.Clear
            End If
            On Error GoTo 0
        Else
            e3App.PutInfo 0, "    ВНИМАНИЕ: Атрибут 'ОД E_Inom' ('" & eInomValue_str & "') пуст или не является числом. Пропускаем обновление KM."
        End If

        If isEInomValueValid Then
            foundMatchingComponent_km = False
            componentName_to_set_km = ""

            For Each componentName_key_km In kmComponentMap.Keys
                rangeValues_km = kmComponentMap.Item(componentName_key_km)
                
                If eInomValue_num >= rangeValues_km(0) And eInomValue_num <= rangeValues_km(1) Then
                    componentName_to_set_km = componentName_key_km
                    foundMatchingComponent_km = True
                    e3App.PutInfo 0, "    Найдено подходящее имя компонента KM: '" & componentName_to_set_km & "' для значения " & eInomValue_num
                    Exit For
                End If
            Next

            If foundMatchingComponent_km Then
                e3App.PutInfo 0, "    Поиск связанных -KM устройств для обновления компонента на: '" & componentName_to_set_km & "'..."
                
                targetDeviceName_KM = "-KM" & oooIndex_str
                Dim kmUpdatedCount : kmUpdatedCount = 0    

                For i = 1 To allDeviceCount
                    deviceObj.SetId(allDeviceIds(i))
                    Dim currentDeviceName : currentDeviceName = deviceObj.GetName()

                    If UCase(currentDeviceName) = UCase(targetDeviceName_KM) Then
                        e3App.PutInfo 0, "      Найдено -KM устройство для обновления: '" & currentDeviceName & "'" & _
                                         " (ID: " & allDeviceIds(i) & ", Текущий компонент: '" & deviceObj.GetComponentName() & "')"
                        
                        On Error Resume Next
                        deviceObj.SetComponentName componentName_to_set_km, COMPONENT_VERSION
                        If Err.Number = 0 Then
                            kmUpdatedCount = kmUpdatedCount + 1
                            e3App.PutInfo 0, "        УСПЕШНО: Компонент KM обновлен на: '" & componentName_to_set_km & "' (Версия: '" & COMPONENT_VERSION & "')."
                        Else
                            e3App.PutInfo 0, "        ОШИБКА при обновлении компонента KM для '" & currentDeviceName & "': " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                Next
                
                If kmUpdatedCount = 0 Then
                    e3App.PutInfo 0, "    Для OOO" & oooIndex_str & " не найдено ни одного -KM устройства для обновления."
                Else
                    e3App.PutInfo 0, "    Всего обновлено " & kmUpdatedCount & " -KM устройств для OOO" & oooIndex_str & "."
                End If
            Else
                e3App.PutInfo 0, "    ВНИМАНИЕ: Для значения 'ОД E_Inom' (" & eInomValue_num & ") не найдено подходящего компонента в таблице соответствия KM. Обновление пропущено."
            End If
        End If    
    Next

    e3App.PutInfo 0, "=== Завершение обновления компонентов -KM ==="

    Set symbolObj = Nothing
    Set deviceObj = Nothing
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
    If Not global_foundOOOIds Is Nothing Then
        Set global_foundOOOIds = Nothing
    End If
    If Not qfComponentMap Is Nothing Then
        Set qfComponentMap = Nothing
    End If
    ' НОВОЕ: Освобождаем объект kmComponentMap
    If Not kmComponentMap Is Nothing Then
        Set kmComponentMap = Nothing
    End If
End Sub

' --- Точка входа в скрипт: запускаем основную процедуру ---
Call Main()