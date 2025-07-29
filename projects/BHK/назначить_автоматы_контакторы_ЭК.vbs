'*******************************************************************************
' Название скрипта: E3_ComponentUpdater_Combined
' Автор: E3.series VBScript Assistant
' Дата: 08.07.2025
' Описание: Скрипт для автоматического обновления имен компонентов для устройств -QF и -KM
'          на основе артикулов, извлеченных из OOO символов, и новой таблицы соответствия.
'          Модифицирован для:
'          1. Отбора OOO символов по атрибуту "ОД D_Proizv3" = "2" ИЛИ "7".
'          2. Умножения значения "ОД E_Inom" на 1.25 для "ОД D_Proizv3" = "2"
'             и на 1.35 для "ОД D_Proizv3" = "7" перед сопоставлением.
'          3. Использования новой таблицы соответствия компонентов для -QF.
'          4. Добавления обновления компонентов для устройств -KM с использованием отдельной таблицы соответствия.
'          5. Исправлена ошибка "Имя было переопределено" путем удаления повторных объявлений Dim.
'          6. Исправлена ошибка "Предполагается наличие инструкции" путем реструктуризации кода без GoTo.
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

' Словарь для хранения соответствий компонента QF и диапазона E_Inom
' Ключ: Имя компонента (String)
' Значение: Массив Double(2) - [МинимальноеЗначение, МаксимальноеЗначение]
Dim qfComponentMap

' Словарь для хранения соответствий компонента KM и диапазона E_Inom
Dim kmComponentMap

' --- Основная процедура скрипта ---
Sub Main()
    ' Инициализация объектов E3.series
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()

    ' Инициализация глобального словаря для хранения найденных OOO символов
    Set global_foundOOOIds = CreateObject("Scripting.Dictionary")
    
    ' Инициализация и заполнение словарей соответствия компонентов
    Set qfComponentMap = CreateObject("Scripting.Dictionary")
    Call PopulateQFComponentMap() ' Процедура для заполнения таблицы соответствий QF

    Set kmComponentMap = CreateObject("Scripting.Dictionary") ' Инициализация нового словаря
    Call PopulateKMComponentMap() ' Процедура для заполнения таблицы соответствий KM

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Поиск OOO символов и связанных с ними устройств ==="

    ' Шаг 1: Находим и фиксируем OOO символы по заданным критериям
    Call FindAndLogOOOSymbols()

    ' Шаг 2: Находим и выводим информацию о связанных устройствах (-QF и -KM)
    ' Обновление: Логирование для -KM также будет показывать компонент "Контактор"
    Call FindAndLogRelatedDevices()

    ' Шаг 3: Обновляем компоненты QF и KM на основе атрибута OOO символа
    ' Обновление: Условное умножение E_Inom в зависимости от D_Proizv3
    Call UpdateComponentsBasedOnOOOAttribute()

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА ==="

    ' Очистка глобальных объектов для освобождения ресурсов
    Call CleanUpGlobalObjects()
End Sub

' --- Процедура для заполнения словаря соответствия компонентов QF ---
Sub PopulateQFComponentMap()
    qfComponentMap.Add "Автомат_3Р_10A_13176DEK", Array(CDbl(0.01), CDbl(10.00))
    qfComponentMap.Add "Автомат_3Р_16A_13177DEK", Array(CDbl(10.01), CDbl(16.00))
    qfComponentMap.Add "Автомат_3Р_20A_13178DEK", Array(CDbl(16.01), CDbl(20.00))
    qfComponentMap.Add "Автомат_3Р_25A_13179DEK", Array(CDbl(20.01), CDbl(25.00))
    qfComponentMap.Add "Автомат_3Р_32A_13180DEK", Array(CDbl(25.01), CDbl(32.00))
    qfComponentMap.Add "Автомат_3Р_40A_13181DEK", Array(CDbl(32.01), CDbl(40.00))
    qfComponentMap.Add "Автомат_3Р_50A_13182DEK", Array(CDbl(40.01), CDbl(50.00))
    qfComponentMap.Add "Автомат_3Р_63A_13183DEK", Array(CDbl(50.01), CDbl(63.00))
    qfComponentMap.Add "Автомат_3Р_80A_13008DEK", Array(CDbl(63.01), CDbl(80.00))
    qfComponentMap.Add "Автомат_3Р_100A_13009DEK", Array(CDbl(80.01), CDbl(100.00))
    qfComponentMap.Add "Автомат_3Р_125A_13027DEK", Array(CDbl(100.01), CDbl(125.00))
    qfComponentMap.Add "Автомат_3Р_160A_22752DEK", Array(CDbl(125.01), CDbl(160.00))
    qfComponentMap.Add "Автомат_3Р_200A_22754DEK", Array(CDbl(160.01), CDbl(200.00))
    qfComponentMap.Add "Автомат_3Р_250A_22756DEK", Array(CDbl(200.01), CDbl(250.00))
    
    e3App.PutInfo 0, "Загружено " & qfComponentMap.Count & " соответствий компонентов для -QF."
End Sub

' --- Процедура для заполнения словаря соответствия компонентов KM ---
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

    e3App.PutInfo 0, "Загружено " & kmComponentMap.Count & " соответствий компонентов для -KM."
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
    
    e3App.PutInfo 0, "Найдено " & allSymbolCount & " символов в проекте. Ищем OOO символы с 'ОД D_Proizv3' = '2' ИЛИ '7'..."

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

            ' ОБНОВЛЕНИЕ: Проверяем, соответствует ли значение атрибута новым критериям ("2" ИЛИ "7")
            If dProizv3Value = "2" Or dProizv3Value = "7" Then
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
        e3App.PutInfo 0, "Не найдено OOO символов со значением атрибута 'ОД D_Proizv3' равным '2' или '7'."
    Else
        e3App.PutInfo 0, "Всего найдено " & foundOOOCount & " OOO символов, соответствующих заданным критериям."
        e3App.PutInfo 0, "ID найденных OOO символов сохранены в глобальном словаре 'global_foundOOOIds'."
    End If

    Set symbol = Nothing ' Освобождаем объект Symbol
End Sub

' --- Процедура для поиска и вывода информации о связанных устройствах (-QF и -KM) ---
Sub FindAndLogRelatedDevices()
    Dim device          ' Объект Device для работы с устройствами
    Dim oooIndex_str    ' Строковое представление числового индекса OOO символа
    Dim targetDeviceName    ' Имя устройства, которое мы ищем (например, "-QF123" или "-KM123")
    Dim allDeviceIds()      ' Массив для хранения идентификаторов всех устройств в проекте
    Dim allDeviceCount      ' Общее количество устройств в проекта
    Dim i                   ' Счетчик цикла для перебора устройств
    Dim currentDeviceName   ' Имя текущего устройства
    Dim componentName       ' Имя компонента текущего устройства

    ' Проверяем, были ли найдены OOO символы на предыдущем шаге
    If global_foundOOOIds.Count = 0 Then
        e3App.PutInfo 0, "Нет зафиксированных OOO символов (D_Proizv3=2 или 7) для поиска связанных устройств."
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
                    e3App.PutInfo 0, "    Найдено -QF устройство: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), но его компонент ('" & componentName & "') не содержит 'Автомат'. Это устройство пропущено."
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
                ' Для -KM устройств, также проверяем, содержит ли компонент "Контактор"
                If InStr(1, LCase(componentName), "контактор") > 0 Then
                    kmFoundCount = kmFoundCount + 1
                    e3App.PutInfo 0, "    Найдено -KM устройство: '" & currentDeviceName & "'" & _
                                 " (ID: " & allDeviceIds(i) & ")" & _
                                 " | Компонент: '" & componentName & "'"
                    foundRelatedDeviceForCurrentOOO = True
                Else
                    e3App.PutInfo 0, "    Найдено -KM устройство: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), но его компонент ('" & componentName & "') не содержит 'Контактор'. Это устройство пропущено."
                End If
            End If
        Next
        
        If kmFoundCount = 0 Then
            e3App.PutInfo 0, "    -KM" & oooIndex_str & " (с компонентом 'Контактор') не найдено ни одного устройства среди всех устройств проекта."
        Else
            e3App.PutInfo 0, "    Всего найдено " & kmFoundCount & " -KM устройств с компонентом 'Контактор' для OOO" & oooIndex_str & "."
        End If

        If Not foundRelatedDeviceForCurrentOOO Then
            e3App.PutInfo 0, "  Для OOO" & oooIndex_str & " не найдено ни одного соответствующего -QF (с компонентом 'Автомат') или -KM (с компонентом 'Контактор') устройства."
        End If
    Next

    e3App.PutInfo 0, "=== Завершение поиска связанных устройств ==="

    Set device = Nothing ' Освобождаем объект Device
End Sub

' --- Процедура для обновления компонентов QF и KM на основе атрибута OOO символа ---
Sub UpdateComponentsBasedOnOOOAttribute()
    Dim symbolObj       ' Объект Symbol для чтения атрибутов OOO
    Dim deviceObj       ' Объект Device для обновления компонентов QF/KM
    Dim oooIndex_str    ' Строковое представление числового индекса OOO символа
    Dim oooSymbolId     ' ID OOO символа
    Dim eInomValue_str  ' Строковое значение атрибута "ОД E_Inom" (исходное)
    Dim eInomValue_num  ' Числовое значение атрибута "ОД E_Inom"
    Dim modifiedEInomValue_num ' Числовое значение "ОД E_Inom" после умножения
    Dim isEInomValueValid ' Флаг для проверки успешности преобразования
    
    Dim targetDeviceName    ' Ожидаемое имя устройства (используется для QF и KM)
    Dim allDeviceIds()      ' Массив ID всех устройств
    Dim allDeviceCount      ' Количество всех устройств
    Dim i                   ' Счетчик цикла

    Dim componentName_to_set ' Имя компонента, который нужно установить
    Dim rangeValues          ' Массив с мин/макс значениями для текущего компонента из словаря
    Dim foundMatchingComponent ' Флаг, указывающий, найден ли подходящий компонент
    Dim componentName_key    ' Переменная для перебора ключей словаря
    Dim trimmedEInomValue_str ' Для обработки строки E_Inom

    ' Переменные для счетчиков обновленных устройств (объявляются один раз)
    Dim qfUpdatedCount
    Dim kmUpdatedCount
    
    ' Переменные для имен текущих устройств (объявляются один раз)
    Dim currentDeviceName
    Dim currentComponentName

    ' Новые переменные для атрибута D_Proizv3 и коэффициента умножения
    Dim dProizv3Value     ' Значение атрибута "ОД D_Proizv3" текущего OOO символа
    Dim multiplicationFactor ' Коэффициент умножения для E_Inom

    ' Константа для версии компонента
    Const COMPONENT_VERSION = "1" ' Версия компонента

    ' Проверяем, были ли найдены OOO символы на предыдущем шаге
    If global_foundOOOIds.Count = 0 Then
        e3App.PutInfo 0, "COMM: Нет зафиксированных OOO символов (D_Proizv3=2 или 7) для обновления компонентов."
        Exit Sub
    End If

    Set symbolObj = job.CreateSymbolObject()
    Set deviceObj = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== Начало обновления компонентов -QF и -KM на основе атрибута 'ОД E_Inom' OOO символов ==="

    ' Проверяем, что словари соответствия компонентов не пусты
    If qfComponentMap.Count = 0 And kmComponentMap.Count = 0 Then
        e3App.PutInfo 0, "ОШИБКА: Словари соответствия компонентов QF и KM пусты. Невозможно обновить компоненты."
        Exit Sub
    End If

    ' Получаем список всех устройств в проекте один раз для эффективности
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' Перебираем каждый зафиксированный OOO символ
    For Each oooIndex_str In global_foundOOOIds.Keys
        oooSymbolId = global_foundOOOIds.Item(oooIndex_str)
        
        ' Устанавливаем OOO символ для чтения атрибута
        symbolObj.SetId(oooSymbolId)
        
        ' Получаем значение атрибута "ОД E_Inom" и "ОД D_Proizv3"
        eInomValue_str = CStr(symbolObj.GetAttributeValue("ОД E_Inom"))    
        dProizv3Value = Trim(CStr(symbolObj.GetAttributeValue("ОД D_Proizv3")))

        e3App.PutInfo 0, "  Обработка OOO" & oooIndex_str & " (ID: " & oooSymbolId & ")"
        e3App.PutInfo 0, "    Исходный атрибут 'ОД E_Inom': '" & eInomValue_str & "'"
        e3App.PutInfo 0, "    Атрибут 'ОД D_Proizv3': '" & dProizv3Value & "'"

        isEInomValueValid = False ' Изначально считаем невалидным
        multiplicationFactor = 0 ' Инициализируем коэффициент

        ' Определяем коэффициент умножения на основе D_Proizv3
        If dProizv3Value = "2" Then
            multiplicationFactor = 1.25
            e3App.PutInfo 0, "    Выбран коэффициент умножения: " & multiplicationFactor & " (для D_Proizv3 = '2')"
        ElseIf dProizv3Value = "7" Then
            multiplicationFactor = 1.35
            e3App.PutInfo 0, "    Выбран коэффициент умножения: " & multiplicationFactor & " (для D_Proizv3 = '7')"
        Else
            e3App.PutInfo 0, "    ВНИМАНИЕ: Неизвестное или неподдерживаемое значение 'ОД D_Proizv3': '" & dProizv3Value & "'. Пропускаем обновление для этого OOO символа."
            ' Здесь нет GoTo. Если multiplicationFactor не установлен, остальной код не выполнится.
        End If
        
        ' Только если multiplicationFactor был установлен (т.е., D_Proizv3 был '2' или '7'),
        ' продолжаем обработку E_Inom и обновление компонентов.
        If multiplicationFactor > 0 Then
            trimmedEInomValue_str = Trim(eInomValue_str)

            If IsNumeric(trimmedEInomValue_str) And Len(trimmedEInomValue_str) > 0 Then
                On Error Resume Next ' Включаем обработку ошибок для CDbl
                eInomValue_num = CDbl(trimmedEInomValue_str)
                If Err.Number = 0 Then
                    isEInomValueValid = True ' Преобразование успешно
                    e3App.PutInfo 0, "    УСПЕШНО: Преобразованное числовое значение 'ОД E_Inom': " & eInomValue_num
                    
                    ' Умножаем на определенный коэффициент
                    modifiedEInomValue_num = eInomValue_num * multiplicationFactor
                    e3App.PutInfo 0, "    Значение 'ОД E_Inom' после умножения на " & multiplicationFactor & ": " & modifiedEInomValue_num
                Else
                    e3App.PutInfo 0, "    ОШИБКА: CDbl не удалось преобразовать строку '" & trimmedEInomValue_str & "' в число (Err: " & Err.Description & ")"
                    Err.Clear ' Очищаем ошибку
                End If
                On Error GoTo 0 ' Отключаем обработку ошибок
            Else
                e3App.PutInfo 0, "    ВНИМАНИЕ: Атрибут 'ОД E_Inom' ('" & eInomValue_str & "') пуст или не является числом. Пропускаем обновление."
            End If

            ' Только если преобразование прошло успешно, ищем подходящий компонент и обновляем
            If isEInomValueValid Then
                ' --- Обновление QF компонентов ---
                e3App.PutInfo 0, "    Поиск и обновление -QF устройств..."
                foundMatchingComponent = False
                componentName_to_set = ""
                qfUpdatedCount = 0 ' Инициализируем счетчик для текущего OOO символа
                
                ' Ищем в словаре QF
                For Each componentName_key In qfComponentMap.Keys
                    rangeValues = qfComponentMap.Item(componentName_key)
                    
                    If modifiedEInomValue_num >= rangeValues(0) And modifiedEInomValue_num <= rangeValues(1) Then
                        componentName_to_set = componentName_key
                        foundMatchingComponent = True
                        e3App.PutInfo 0, "      Найдено подходящее имя компонента QF: '" & componentName_to_set & "'"
                        Exit For
                    End If
                Next

                If foundMatchingComponent Then
                    targetDeviceName = "-QF" & oooIndex_str
                    
                    For i = 1 To allDeviceCount
                        deviceObj.SetId(allDeviceIds(i))
                        currentDeviceName = deviceObj.GetName()
                        currentComponentName = deviceObj.GetComponentName()

                        If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                            If InStr(1, LCase(currentComponentName), "автомат") > 0 Then
                                On Error Resume Next
                                deviceObj.SetComponentName componentName_to_set, COMPONENT_VERSION
                                If Err.Number = 0 Then
                                    qfUpdatedCount = qfUpdatedCount + 1
                                    e3App.PutInfo 0, "        УСПЕШНО: Компонент -QF '" & currentDeviceName & "' обновлен на: '" & componentName_to_set & "'."
                                Else
                                    e3App.PutInfo 0, "        ОШИБКА при обновлении компонента QF для '" & currentDeviceName & "': " & Err.Description
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            Else
                                e3App.PutInfo 0, "      Найдено -QF устройство: '" & currentDeviceName & "', но его компонент ('" & currentComponentName & "') не содержит 'Автомат'. Пропущено обновление."
                            End If
                        End If
                    Next
                    If qfUpdatedCount = 0 Then
                        e3App.PutInfo 0, "    Для OOO" & oooIndex_str & " не найдено ни одного -QF устройства с компонентом 'Автомат' для обновления."
                    Else
                        e3App.PutInfo 0, "    Всего обновлено " & qfUpdatedCount & " -QF устройств для OOO" & oooIndex_str & "."
                    End If
                Else
                    e3App.PutInfo 0, "    ВНИМАНИЕ: Для модифицированного значения " & modifiedEInomValue_num & " не найдено подходящего компонента QF в таблице соответствия. Обновление QF пропущено."
                End If


                ' --- Обновление KM компонентов ---
                e3App.PutInfo 0, "    Поиск и обновление -KM устройств..."
                foundMatchingComponent = False
                componentName_to_set = ""
                kmUpdatedCount = 0 ' Инициализируем счетчик для текущего OOO символа

                ' Ищем в словаре KM
                For Each componentName_key In kmComponentMap.Keys
                    rangeValues = kmComponentMap.Item(componentName_key)
                    
                    If modifiedEInomValue_num >= rangeValues(0) And modifiedEInomValue_num <= rangeValues(1) Then
                        componentName_to_set = componentName_key
                        foundMatchingComponent = True
                        e3App.PutInfo 0, "      Найдено подходящее имя компонента KM: '" & componentName_to_set & "'"
                        Exit For
                    End If
                Next

                If foundMatchingComponent Then
                    targetDeviceName = "-KM" & oooIndex_str
                    
                    For i = 1 To allDeviceCount
                        deviceObj.SetId(allDeviceIds(i))
                        currentDeviceName = deviceObj.GetName()
                        currentComponentName = deviceObj.GetComponentName()

                        If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                            ' Дополнительная проверка, что компонент KM содержит "Контактор"
                            If InStr(1, LCase(currentComponentName), "контактор") > 0 Then
                                On Error Resume Next
                                deviceObj.SetComponentName componentName_to_set, COMPONENT_VERSION
                                If Err.Number = 0 Then
                                    kmUpdatedCount = kmUpdatedCount + 1
                                    e3App.PutInfo 0, "        УСПЕШНО: Компонент -KM '" & currentDeviceName & "' обновлен на: '" & componentName_to_set & "'."
                                Else
                                    e3App.PutInfo 0, "        ОШИБКА при обновлении компонента KM для '" & currentDeviceName & "': " & Err.Description
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            Else
                                e3App.PutInfo 0, "      Найдено -KM устройство: '" & currentDeviceName & "', но его компонент ('" & currentComponentName & "') не содержит 'Контактор'. Пропущено обновление."
                            End If
                        End If
                    Next
                    If kmUpdatedCount = 0 Then
                        e3App.PutInfo 0, "    Для OOO" & oooIndex_str & " не найдено ни одного -KM устройства с компонентом 'Контактор' для обновления."
                    Else
                        e3App.PutInfo 0, "    Всего обновлено " & kmUpdatedCount & " -KM устройств для OOO" & oooIndex_str & "."
                    End If
                Else
                    e3App.PutInfo 0, "    ВНИМАНИЕ: Для модифицированного значения " & modifiedEInomValue_num & " не найдено подходящего компонента KM в таблице соответствия. Обновление KM пропущено."
                End If
            End If ' End If isEInomValueValid Then
        End If ' End If multiplicationFactor > 0 Then
    Next ' Продолжаем к следующему OOO символу

    e3App.PutInfo 0, "=== Завершение обновления компонентов -QF и -KM ==="

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
    If Not global_foundOOOIds Is Nothing Then
        Set global_foundOOOIds = Nothing
    End If
    ' Освобождаем объекты словарей компонентов
    If Not qfComponentMap Is Nothing Then
        Set qfComponentMap = Nothing
    End If
    If Not kmComponentMap Is Nothing Then
        Set kmComponentMap = Nothing
    End If
End Sub

' --- Точка входа в скрипт: запускаем основную процедуру ---
Call Main()