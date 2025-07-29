Option Explicit

' === Процедура 1 === Поиск устройств по выделенным символам
' Эта процедура в основном используется для определения наличия символа OOO
' и для вывода диагностической информации о выделенных символах.
Sub FindDevicesBySelectedSymbols(foundDevices, ByRef hasOOOSymbol)
    Dim e3App, job, device, symbol
    Dim selectedSymbolIds(), selectedCount
    Dim deviceIds(), deviceCount
    Dim symbolIds(), result
    Dim selectedSymbolId, symbolRealId, currentSymbolId
    Dim foundDeviceId, symbolName
    Dim i, j, s

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set device = job.CreateDeviceObject()
    Set symbol = job.CreateSymbolObject()

    hasOOOSymbol = False ' Инициализируем флаг OOO символа

    selectedCount = job.GetSelectedSymbolIds(selectedSymbolIds)
    If selectedCount = 0 Then
        e3App.PutInfo 0, "Нет выделенных символов на схеме."
        ' В этом случае скрипт продолжит работу, чтобы запросить имена QF/QS и KM,
        ' но OOO символ не будет найден, и атрибуты не будут записаны.
    Else
        e3App.PutInfo 0, "=== ОБРАБОТКА ВЫДЕЛЕННЫХ СИМВОЛОВ ==="
        For s = 1 To selectedCount
            selectedSymbolId = selectedSymbolIds(s)
            symbol.SetId(selectedSymbolId)
            symbolRealId = symbol.GetId()
            symbolName = symbol.GetName()

            If LCase(Left(symbolName, 3)) = "ooo" Then
                e3App.PutInfo 0, "Символ OOO: " & symbolName & " (ID: " & symbolRealId & ") — найден."
                hasOOOSymbol = True ' Устанавливаем флаг
            Else
                ' Здесь логика поиска устройства по символу сохранена для диагностических целей,
                ' но для поиска QF/QS и KM устройств будет использоваться прямой ввод пользователя.
                foundDeviceId = 0
                deviceCount = job.GetAllDeviceIds(deviceIds) ' Получаем все устройства для поиска символов

                If deviceCount > 0 Then
                    For i = 1 To deviceCount
                        device.SetId(deviceIds(i))
                        result = device.GetSymbolIds(symbolIds, 3) ' Получаем символы, связанные с устройством

                        If result > 0 Then
                            For j = 1 To result
                                symbol.SetId(symbolIds(j))
                                currentSymbolId = symbol.GetId()

                                If currentSymbolId = symbolRealId Then
                                    foundDeviceId = deviceIds(i)
                                    Exit For ' Символ найден для этого устройства
                                    End If
                            Next
                        End If
                        If foundDeviceId <> 0 Then Exit For ' Устройство найдено для текущего символа
                    Next
                End If ' Закрывающий If для If deviceCount > 0 Then

                If foundDeviceId <> 0 Then
                    If Not foundDevices.Exists(CStr(foundDeviceId)) Then
                        device.SetId(foundDeviceId)
                        e3App.PutInfo 0, "Найдено устройство для символа: " & device.GetName() & " (ID: " & foundDeviceId & ")"
                        foundDevices.Add CStr(foundDeviceId), True ' Добавляем в коллекцию найденных устройств
                    End If
                Else
                    e3App.PutInfo 0, "Устройство для символа '" & symbolName & "' (ID: " & symbolRealId & ") не найдено."
                End If
            End If
        Next
        e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ ОБРАБОТКИ ВЫДЕЛЕННЫХ СИМВОЛОВ ==="
    End If
    
    e3App.PutInfo 0, "=== ИТОГО найдено устройств по выделенным символам: " & foundDevices.Count & " ==="
    e3App.PutInfo 0, "=== Символ OOO найден: " & hasOOOSymbol & " ==="

    Set symbol = Nothing
    Set device = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === Процедура 2 === Основная процедура обработки атрибутов устройств (QF/SF и KM)
Sub ProcessAndWriteDeviceAttributes(hasOOOSymbol)
    Dim e3App, job, device
    Dim deviceIds(), deviceCount ' Переменные для всех устройств в проекте
    Dim i
    Dim userDeviceName ' Переменная для ввода пользователя (общее название для QF/SF и KM)
    Dim currentComponentName ' Для временного хранения имени компонента при итерации
    Dim key ' Объявление переменной key
    Dim deviceName ' Объявление переменной deviceName

    ' --- Переменные для атрибутов устройства QF/SF "Автомат" ---
    Dim selectedAutomatDeviceId : selectedAutomatDeviceId = 0
    Dim automatTechDesc : automatTechDesc = ""
    Dim automatCompTitle : automatCompTitle = ""
    Dim automatCompSupplier : automatCompSupplier = ""

    ' --- Переменные для атрибутов устройства QF/SF "Контакт" ---
    Dim selectedContactDeviceId : selectedContactDeviceId = 0
    Dim contactCompName : contactCompName = ""

    ' --- Переменные для атрибутов устройства KM "Контактор" ---
    Dim selectedContactorDeviceId : selectedContactorDeviceId = 0
    Dim kmCompTitle : kmCompTitle = ""
    Dim kmCompSupplier : kmCompSupplier = ""
    Dim kmCompCurrent : kmCompCurrent = "" ' Для атрибута "(Класс) Ток"

    ' Используем словари для хранения всех найденных QF/SF и KM устройств по имени
    Dim matchingQfDevices
    Set matchingQfDevices = CreateObject("Scripting.Dictionary")
    Dim matchingKmDevices
    Set matchingKmDevices = CreateObject("Scripting.Dictionary")
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set device = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== ЗАПУСК ПРОЦЕДУРЫ ОБРАБОТКИ АТРИБУТОВ ==="
    
    ' Получаем все ID устройств в проекте ОДИН РАЗ
    deviceCount = job.GetAllDeviceIds(deviceIds)
    If deviceCount = 0 Then
        e3App.PutInfo 0, "ОШИБКА: В проекте нет устройств для поиска. Скрипт завершен."
        Set device = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    ' --- Обработка QF/SF устройств (обязательно) ---
    e3App.PutInfo 0, "--- ОБРАБОТКА УСТРОЙСТВ QF/SF ---"
    userDeviceName = InputBox("Введите полное имя устройства QF/SF для обработки:", "Имя устройства QF/SF", "-QF")

    If Trim(userDeviceName) = "" Then
        e3App.PutInfo 0, "Действие отменено пользователем или имя устройства QF/SF не введено. Пропуск обработки QF/SF."
    Else
        e3App.PutInfo 0, "Ищем QF/SF устройство по имени: '" & userDeviceName & "' во всем проекте..."
        For i = 1 To deviceCount
            device.SetId(deviceIds(i))
            deviceName = device.GetName()

            If LCase(deviceName) = LCase(userDeviceName) Then
                matchingQfDevices.Add CStr(deviceIds(i)), deviceName
                e3App.PutInfo 0, "  НАЙДЕНО QF/SF-устройство: " & deviceName & " (ID: " & deviceIds(i) & ")"
            End If
        Next

        If matchingQfDevices.Count = 0 Then
            e3App.PutInfo 0, "Устройство QF/SF с именем '" & userDeviceName & "' НЕ НАЙДЕНО во всем проекте."
        Else
            e3App.PutInfo 0, "Проверяем найденные QF/SF устройства на наличие 'Автомат' и 'Контакт' в имени компонента..."
            
            ' Ищем устройство с компонентом "Автомат" среди QF/SF
            For Each key In matchingQfDevices.Keys
                device.SetId CLng(key)
                currentComponentName = device.GetComponentName()
                
                If InStr(1, LCase(currentComponentName), "автомат") > 0 Then
                    selectedAutomatDeviceId = CLng(key)
                    automatTechDesc = device.GetAttributeValue("Тех. описание 1")
                    automatCompTitle = device.GetComponentAttributeValue("Наименование")
                    automatCompSupplier = device.GetComponentAttributeValue("Поставщик")
                    e3App.PutInfo 0, "  Выбрано для обработки (Автомат): " & device.GetName() & " (ID: " & selectedAutomatDeviceId & ")"
                    ' Продолжаем, чтобы найти "Контакт"
                End If
            Next

            ' Ищем устройство с компонентом "Контакт" среди QF/SF
            For Each key In matchingQfDevices.Keys
                device.SetId CLng(key)
                currentComponentName = device.GetComponentName()

                If InStr(1, LCase(currentComponentName), "контакт") > 0 Then
                    selectedContactDeviceId = CLng(key)
                    contactCompName = device.GetComponentAttributeValue("Наименование")
                    e3App.PutInfo 0, "  Выбрано для обработки (Контакт): " & device.GetName() & " (ID: " & selectedContactDeviceId & ") (Наименование компонента: '" & contactCompName & "')"
                    ' Продолжаем
                End If
            Next
            
            If selectedAutomatDeviceId = 0 Then
                e3App.PutInfo 0, "  ПРЕДУПРЕЖДЕНИЕ: Для QF/SF с именем '" & userDeviceName & "' компонент 'Автомат' не найден."
            End If
            If selectedContactDeviceId = 0 Then
                e3App.PutInfo 0, "  ПРЕДУПРЕЖДЕНИЕ: Для QF/SF с именем '" & userDeviceName & "' компонент 'Контакт' не найден."
            End If
        End If
    End If

    ' --- Обработка KM устройств (опционально) ---
    e3App.PutInfo 0, "--- ОБРАБОТКА УСТРОЙСТВ KM ---"
    ' Важно: userDeviceName переиспользуется, поэтому значение от QF/SF здесь сбрасывается.
    userDeviceName = InputBox("Введите полное имя устройства KM для обработки (оставьте пустым, если КМ нет):", "Имя устройства KM (опционально)", "-KM")

    If Trim(userDeviceName) = "" Then
        e3App.PutInfo 0, "Действие отменено пользователем или имя устройства KM не введено. Пропуск обработки KM."
    Else
        e3App.PutInfo 0, "Ищем KM устройство по имени: '" & userDeviceName & "' во всем проекте..."
        For i = 1 To deviceCount
            device.SetId(deviceIds(i))
            deviceName = device.GetName()
            If LCase(deviceName) = LCase(userDeviceName) Then
                matchingKmDevices.Add CStr(deviceIds(i)), deviceName
                e3App.PutInfo 0, "  НАЙДЕНО KM-устройство: " & deviceName & " (ID: " & deviceIds(i) & ")"
            End If
        Next

        If matchingKmDevices.Count = 0 Then
            e3App.PutInfo 0, "Устройство KM с именем '" & userDeviceName & "' НЕ НАЙДЕНО во всем проекте."
        Else
            e3App.PutInfo 0, "Проверяем найденные KM устройства на наличие 'Контактор' в имени компонента..."
            For Each key In matchingKmDevices.Keys
                device.SetId CLng(key)
                currentComponentName = device.GetComponentName()
                If InStr(1, LCase(currentComponentName), "контактор") > 0 Then
                    selectedContactorDeviceId = CLng(key)
                    kmCompTitle = device.GetComponentAttributeValue("Наименование")
                    kmCompSupplier = device.GetComponentAttributeValue("Поставщик")
                    kmCompCurrent = device.GetComponentAttributeValue("(Класс) Ток") ' Читаем атрибут "(Класс) Ток"
                    e3App.PutInfo 0, "  Выбрано для обработки (Контактор): " & device.GetName() & " (ID: " & selectedContactorDeviceId & ")"
                    Exit For ' Берем первое найденное устройство "Контактор"
                End If
            Next
            If selectedContactorDeviceId = 0 Then
                 e3App.PutInfo 0, "ПРЕДУПРЕЖДЕНИЕ: Среди найденных KM устройств с именем '" & userDeviceName & "' не найдено ни одного с компонентом, содержащим 'Контактор'."
            End If
        End If
    End If

    ' --- Сводка и запись атрибутов в OOO символ ---
    If hasOOOSymbol Then
        ' Вызываем запись только если хоть одно из устройств было обработано
        If (selectedAutomatDeviceId > 0 Or selectedContactDeviceId > 0) Or selectedContactorDeviceId > 0 Then
            e3App.PutInfo 0, "--- Сводка собранных атрибутов для записи в OOO символ ---"
            If selectedAutomatDeviceId > 0 Then
                e3App.PutInfo 0, "  Атрибуты от QF/SF (Автомат) будут записаны."
            Else
                e3App.PutInfo 0, "  Атрибуты QF/SF (Автомат) отсутствуют (будут пустыми)."
            End If
            If selectedContactDeviceId > 0 Then
                e3App.PutInfo 0, "  Атрибуты от QF/SF (Контакт) будут записаны."
            Else
                e3App.PutInfo 0, "  Атрибуты QF/SF (Контакт) отсутствуют (будут пустыми)."
            End If
            If selectedContactorDeviceId > 0 Then
                e3App.PutInfo 0, "  Атрибуты от KM (Контактор) будут записаны."
            Else
                e3App.PutInfo 0, "  Атрибуты KM (Контактор) отсутствуют (будут пустыми)."
            End If
            e3App.PutInfo 0, "-----------------------------------------------------"
            
            Call WriteAttributesToOOOSymbol(automatTechDesc, automatCompTitle, automatCompSupplier, _
                                            contactCompName, _
                                            kmCompTitle, kmCompSupplier, kmCompCurrent)
        Else
            e3App.PutInfo 0, "ПРЕДУПРЕЖДЕНИЕ: Не найдено ни одного подходящего устройства (QF/SF Автомат/Контакт или KM Контактор) для записи атрибутов в OOO символ."
        End If
    Else
        e3App.PutInfo 0, "Символ OOO не был найден среди выделенных, атрибуты не записаны."
    End If

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ ПРОЦЕДУРЫ ОБРАБОТКИ АТРИБУТОВ ==="

    Set device = Nothing
    Set job = Nothing
    Set e3App = Nothing
    Set matchingQfDevices = Nothing
    Set matchingKmDevices = Nothing
End Sub

' === ФУНКЦИЯ === Извлечение значения "kA" из текста
Function ExtractkAValue(ByVal inputText)
    Dim re, matches
    Set re = New RegExp
    ' Регулярное выражение для поиска чисел, за которыми сразу следует "kA" (например, "10kA", "100kA")
    re.Pattern = "\b(\d+kA)\b" 
    re.IgnoreCase = True ' Игнорировать регистр ("kA" или "KA")
    re.Global = False    ' Найти только первое совпадение

    Set matches = re.Execute(inputText)

    If matches.Count > 0 Then
        ExtractkAValue = matches.Item(0).SubMatches.Item(0) ' Получаем захваченную группу (например, "10kA")
    Else
        ExtractkAValue = "" ' Если не найдено, возвращаем пустую строку
    End If

    Set re = Nothing
End Function

' === Процедура 3 === Запись атрибутов QF/SF и KM в символ OOO
Sub WriteAttributesToOOOSymbol(ByVal automatTechDesc, ByVal automatCompTitle, ByVal automatCompSupplier, _
                               ByVal contactComponentNameForOOO, _
                               ByVal kmCompTitle, ByVal kmCompSupplier, ByVal kmCompCurrent)
    Dim e3App, job, symbol
    Dim selectedSymbolIds(), selectedCount
    Dim selectedSymbolId, symbolName
    Dim s

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()

    e3App.PutInfo 0, "=== ЗАПИСЬ атрибутов в символ OOO ==="

    selectedCount = job.GetSelectedSymbolIds(selectedSymbolIds)
    If selectedCount = 0 Then
        e3App.PutInfo 0, "ОШИБКА: Нет выделенных символов для записи атрибутов в OOO. Возможно, символ OOO был удален или не выделен."
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    
    Dim oooSymbolFoundForWriting : oooSymbolFoundForWriting = False

    For s = 1 To selectedCount
        selectedSymbolId = selectedSymbolIds(s)
        symbol.SetId(selectedSymbolId)
        symbolName = symbol.GetName()

        If LCase(Left(symbolName, 3)) = "ooo" Then
            e3App.PutInfo 0, "Найден символ OOO для записи: " & symbolName & " (ID: " & selectedSymbolId & ")"
            
            ' --- Записываем атрибуты из QF/SF устройства (предполагается 'Автомат') ---
            If Len("" & automatTechDesc) > 0 Then
                symbol.SetAttributeValue "ОД V_Inom", automatTechDesc
                e3App.PutInfo 0, "Записано в ОД V_Inom (из Тех. описание 1 Автомата QF/SF): " & automatTechDesc
            Else
                e3App.PutInfo 0, "Атрибут 'Тех. описание 1' Автомата QF/SF пуст, нечего записывать в ОД V_Inom."
            End If
            
            If Len("" & automatCompTitle) > 0 Then
                symbol.SetAttributeValue "ОД V_Type", automatCompTitle
                e3App.PutInfo 0, "Записано в ОД V_Type (из Наименования компонента Автомата QF/SF): " & automatCompTitle
                
                ' Извлечение и запись значения "kA"
                Dim extractedkAValue
                extractedkAValue = ExtractkAValue(automatCompTitle) ' Используем новую функцию
                If Len("" & extractedkAValue) > 0 Then
                    symbol.SetAttributeValue "ОД V_Icu", extractedkAValue
                    e3App.PutInfo 0, "Записано в ОД V_Icu (из Наименования компонента Автомата QF/SF): " & extractedkAValue
                Else
                    e3App.PutInfo 0, "Значение 'kA' не найдено в Наименовании компонента Автомата QF/SF, нечего записывать в ОД V_Icu."
                End If

            Else
                e3App.PutInfo 0, "Атрибут 'Наименование' компонента Автомата QF/SF пуст, нечего записывать в ОД V_Type и ОД V_Icu."
            End If
            
            If Len("" & automatCompSupplier) > 0 Then
                symbol.SetAttributeValue "ОД V_Proizv", automatCompSupplier
                e3App.PutInfo 0, "Записано в ОД V_Proizv (из Поставщика компонента Автомата QF/SF): " & automatCompSupplier
            Else
                e3App.PutInfo 0, "Атрибут 'Поставщик' компонента Автомата QF/SF пуст, нечего записывать в ОД V_Proizv."
            End If
            
            ' --- Записываем атрибут для компонента "Контакт" ---
            If Len("" & contactComponentNameForOOO) > 0 Then
                symbol.SetAttributeValue "ОД V_Dop ystr", contactComponentNameForOOO
                e3App.PutInfo 0, "Записано в ОД V_Dop ystr (из Наименования компонента 'Контакт' QF/SF): " & contactComponentNameForOOO
            Else
                e3App.PutInfo 0, "Атрибут 'Наименование' компонента 'Контакт' QF/SF пуст или не найден, нечего записывать в ОД V_Dop ystr."
            End If

            ' --- Записываем атрибуты из KM устройства (предполагается 'Контактор') ---
            If Len("" & kmCompTitle) > 0 Then
                symbol.SetAttributeValue "ОД K_Type", kmCompTitle
                e3App.PutInfo 0, "Записано в ОД K_Type (из Наименования компонента Контактора KM): " & kmCompTitle
            Else
                e3App.PutInfo 0, "Атрибут 'Наименование' компонента Контактора KM пуст, нечего записывать в ОД K_Type."
            End If

            If Len("" & kmCompSupplier) > 0 Then
                symbol.SetAttributeValue "ОД K_Proizv", kmCompSupplier
                e3App.PutInfo 0, "Записано в ОД K_Proizv (из Поставщика компонента Контактора KM): " & kmCompSupplier
            Else
                e3App.PutInfo 0, "Атрибут 'Поставщик' компонента Контактора KM пуст, нечего записывать в ОД K_Proizv."
            End If

            If Len("" & kmCompCurrent) > 0 Then
                symbol.SetAttributeValue "ОД K_Inom", kmCompCurrent
                e3App.PutInfo 0, "Записано в ОД K_Inom (из (Класс) Ток компонента Контактора KM): " & kmCompCurrent
            Else
                e3App.PutInfo 0, "Атрибут '(Класс) Ток' компонента Контактора KM пуст, нечего записывать в ОД K_Inom."
            End If


            e3App.PutInfo 0, "Атрибуты успешно записаны в символ OOO."
            oooSymbolFoundForWriting = True
            Exit For ' Записали в первый найденный OOO символ и выходим
        End If
    Next

    If Not oooSymbolFoundForWriting Then
        e3App.PutInfo 0, "Не найден символ OOO среди выделенных для записи атрибутов."
    End If

    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub


' === Основной запуск ===
Dim foundDevices, e3App, hasOOOSymbol
Set foundDevices = CreateObject("Scripting.Dictionary") ' Эта коллекция теперь используется в основном для диагностики
Set e3App = CreateObject("CT.Application")

e3App.PutInfo 0, "=== СТАРТ СКРИПТА ==="
Call FindDevicesBySelectedSymbols(foundDevices, hasOOOSymbol) ' Определяем, есть ли OOO символ среди выделенных
Call ProcessAndWriteDeviceAttributes(hasOOOSymbol) ' Теперь передаем только флаг OOO символа
e3App.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="

Set foundDevices = Nothing
Set e3App = Nothing
