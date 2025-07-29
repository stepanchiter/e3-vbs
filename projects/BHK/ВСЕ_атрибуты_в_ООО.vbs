'*******************************************************************************
' Название скрипта: E3_QF_TechDescUpdater_V_Devices
' Автор: E3.series VBScript Assistant
' Дата: 08.07.2025
' Описание: Записывает атрибуты аппаратов в ООО. Скрипт для автоматической записи значения в атрибут "Тех. описание 3"
'           для устройств QF, исключая те, в компоненте которых содержится "ДИФ".
'           Добавлена обработка устройств с префиксом -V и запись их имени компонента
'           в атрибут "ОД D_Proizv2" соответствующего символа OOO.
'*******************************************************************************

Option Explicit

' === Функция === Извлечение номера из имени символа/устройства
Function ExtractNumber(ByVal itemName)
    Dim re, matches
    Set re = New RegExp
    ' Ищем число в конце строки после символов (например, OOO1, -QF1, -KM1, -KL1, -V1)
    re.Pattern = "(\d+)$"
    re.Global = False
    
    Set matches = re.Execute(itemName)
    
    If matches.Count > 0 Then
        ExtractNumber = CInt(matches.Item(0).Value)
    Else
        ExtractNumber = 0 ' Если номер не найден
    End If
    
    Set re = Nothing
End Function

' === Функция === Извлечение значения "kA" из текста
Function ExtractkAValue(ByVal inputText)
    Dim re, matches
    Set re = New RegExp
    ' Регулярное выражение для поиска чисел, за которыми сразу следует "kA"
    re.Pattern = "\b(\d+kA)\b" 
    re.IgnoreCase = True
    re.Global = False
    
    Set matches = re.Execute(inputText)
    
    If matches.Count > 0 Then
        ExtractkAValue = matches.Item(0).SubMatches.Item(0)
    Else
        ExtractkAValue = ""
    End If
    
    Set re = Nothing
End Function

' === Функция === Получение значения атрибута ОД E_Inom из символа OOO
Function GetOOOAttributeEInom(ByVal e3AppObj, ByVal jobObj, ByVal oooSymbolId)
    Dim symbol, attributeValue
    
    Set symbol = jobObj.CreateSymbolObject()
    
    symbol.SetId(oooSymbolId)
    attributeValue = symbol.GetAttributeValue("ОД E_Inom")
    
    GetOOOAttributeEInom = attributeValue
    
    Set symbol = Nothing
End Function

' === Процедура === Запись значения в атрибут "Тех. описание 3" QF устройства
' Модифицирована для исключения устройств с "ДИФ" в имени компонента.
Sub WriteToQFDeviceTechDesc3(ByVal e3AppObj, ByVal jobObj, ByVal deviceId, ByVal value)
    Dim device
    Dim componentName
    
    Set device = jobObj.CreateDeviceObject()
    
    device.SetId(deviceId)
    
    ' Получаем имя компонента для проверки
    componentName = device.GetComponentName()
    
    ' Проверяем, не содержит ли имя компонента "ДИФ"
    If InStr(1, LCase(componentName), "диф") = 0 Then
        device.SetAttributeValue "Тех. описание 3", value
        e3AppObj.PutInfo 0, "Записано в QF устройство (ID: " & deviceId & ", Компонент: " & componentName & ") Тех. описание 3: " & value
    Else
        e3AppObj.PutInfo 0, "Пропущено QF устройство (ID: " & deviceId & ", Компонент: " & componentName & "): содержит 'ДИФ' в имени компонента."
    End If
    
    Set device = Nothing
End Sub

' === Процедура === Поиск всех символов OOO в проекте
Sub FindAllOOOSymbols(ByVal e3AppObj, ByVal jobObj, ByRef oooSymbols)
    Dim symbol
    Dim symbolIds(), symbolCount
    Dim i, symbolName, symbolNumber
    
    Set symbol = jobObj.CreateSymbolObject()
    
    e3AppObj.PutInfo 0, "=== ПОИСК ВСЕХ СИМВОЛОВ OOO В ПРОЕКТЕ ==="
    
    symbolCount = jobObj.GetSymbolIds(symbolIds)
    If symbolCount = 0 Then
        e3AppObj.PutInfo 0, "В проекте не найдено символов."
        Set symbol = Nothing
        Exit Sub
    End If
    
    For i = 1 To symbolCount
        symbol.SetId(symbolIds(i))
        symbolName = symbol.GetName()
        
        If LCase(Left(symbolName, 3)) = "ooo" Then
            symbolNumber = ExtractNumber(symbolName)
            If symbolNumber > 0 Then
                oooSymbols.Add symbolNumber, symbolIds(i)
                e3AppObj.PutInfo 0, "Найден символ OOO: " & symbolName & " (номер: " & symbolNumber & ", ID: " & symbolIds(i) & ")"
            Else
                e3AppObj.PutInfo 0, "Символ OOO найден, но номер не определен: " & symbolName
            End If
        End If
    Next
    
    e3AppObj.PutInfo 0, "Всего найдено символов OOO с номерами: " & oooSymbols.Count
    
    Set symbol = Nothing
End Sub

' === Процедура === Поиск всех устройств QF, KM, KL, V в проекте
Sub FindAllDevices(ByVal e3AppObj, ByVal jobObj, ByRef qfDevices, ByRef kmDevices, ByRef klDevices, ByRef vDevices)
    Dim device
    Dim deviceIds(), deviceCount
    Dim i, deviceName, deviceNumber
    
    Set device = jobObj.CreateDeviceObject()
    
    e3AppObj.PutInfo 0, "=== ПОИСК ВСЕХ УСТРОЙСТВ QF, KM, KL, V В ПРОЕКТЕ ==="
    
    deviceCount = jobObj.GetAllDeviceIds(deviceIds)
    If deviceCount = 0 Then
        e3AppObj.PutInfo 0, "В проекте не найдено устройств."
        Set device = Nothing
        Exit Sub
    End If
    
    For i = 1 To deviceCount
        device.SetId(deviceIds(i))
        deviceName = device.GetName()
        
        ' Поиск устройств QF
        If InStr(1, LCase(deviceName), "-qf") > 0 Then
            deviceNumber = ExtractNumber(deviceName)
            If deviceNumber > 0 Then
                If Not qfDevices.Exists(deviceNumber) Then
                    qfDevices.Add deviceNumber, CreateObject("Scripting.Dictionary")
                End If
                qfDevices.Item(deviceNumber).Add CStr(deviceIds(i)), deviceName
                e3AppObj.PutInfo 0, "Найдено QF устройство: " & deviceName & " (номер: " & deviceNumber & ", ID: " & deviceIds(i) & ")"
            End If
        End If
        
        ' Поиск устройств KM
        If InStr(1, LCase(deviceName), "-km") > 0 Then
            deviceNumber = ExtractNumber(deviceName)
            If deviceNumber > 0 Then
                If Not kmDevices.Exists(deviceNumber) Then
                    kmDevices.Add deviceNumber, CreateObject("Scripting.Dictionary")
                End If
                kmDevices.Item(deviceNumber).Add CStr(deviceIds(i)), deviceName
                e3AppObj.PutInfo 0, "Найдено KM устройство: " & deviceName & " (номер: " & deviceNumber & ", ID: " & deviceIds(i) & ")"
            End If
        End If
        
        ' Поиск устройств KL
        If InStr(1, LCase(deviceName), "-kl") > 0 Then
            deviceNumber = ExtractNumber(deviceName)
            If deviceNumber > 0 Then
                If Not klDevices.Exists(deviceNumber) Then
                    klDevices.Add deviceNumber, CreateObject("Scripting.Dictionary")
                End If
                klDevices.Item(deviceNumber).Add CStr(deviceIds(i)), deviceName
                e3AppObj.PutInfo 0, "Найдено KL устройство: " & deviceName & " (номер: " & deviceNumber & ", ID: " & deviceIds(i) & ")"
            End If
        End If

        ' Поиск устройств V
        If InStr(1, LCase(deviceName), "-v") > 0 Then
            deviceNumber = ExtractNumber(deviceName)
            If deviceNumber > 0 Then
                If Not vDevices.Exists(deviceNumber) Then
                    vDevices.Add deviceNumber, CreateObject("Scripting.Dictionary")
                End If
                vDevices.Item(deviceNumber).Add CStr(deviceIds(i)), deviceName
                e3AppObj.PutInfo 0, "Найдено V устройство: " & deviceName & " (номер: " & deviceNumber & ", ID: " & deviceIds(i) & ")"
            End If
        End If
    Next
    
    e3AppObj.PutInfo 0, "Найдено QF устройств с номерами: " & qfDevices.Count
    e3AppObj.PutInfo 0, "Найдено KM устройств с номерами: " & kmDevices.Count
    e3AppObj.PutInfo 0, "Найдено KL устройств с номерами: " & klDevices.Count
    e3AppObj.PutInfo 0, "Найдено V устройств с номерами: " & vDevices.Count
    
    Set device = Nothing
End Sub

' === Процедура === Получение атрибутов устройства
' Модифицирована для возврата имени компонента
Sub GetDeviceAttributes(ByVal jobObj, ByVal deviceId, ByRef techDesc, ByRef compTitle, ByRef compSupplier, ByRef compCurrent, ByRef compName)
    Dim device
    
    Set device = jobObj.CreateDeviceObject()
    
    device.SetId(deviceId)
    
    ' techDesc используется только для QF "Тех. описание 1"
    If Not IsEmpty(techDesc) Then techDesc = device.GetAttributeValue("Тех. описание 1")
    
    ' compTitle используется для QF, KM, KL, V "Наименование"
    If Not IsEmpty(compTitle) Then compTitle = device.GetComponentAttributeValue("Наименование")
    
    ' compSupplier используется только для QF, KM "Поставщик"
    If Not IsEmpty(compSupplier) Then compSupplier = device.GetComponentAttributeValue("Поставщик")
    
    ' compCurrent используется только для KM "(Класс) Ток"
    If Not IsEmpty(compCurrent) Then compCurrent = device.GetComponentAttributeValue("(Класс) Ток")
    
    ' compName используется для всех, чтобы получить имя компонента
    compName = device.GetComponentName()
    
    Set device = Nothing
End Sub

' === Процедура === Запись атрибутов в символ OOO
Sub WriteAttributesToOOOSymbol(ByVal e3AppObj, ByVal jobObj, ByVal oooSymbolId, ByVal number, _
                                 ByVal qfTechDesc, ByVal qfCompTitle, ByVal qfCompSupplier, ByVal qfContactCompName, _
                                 ByVal kmCompTitle, ByVal kmCompSupplier, ByVal kmCompCurrent, _
                                 ByVal klCompTitle, ByVal vCompName)
    Dim symbol
    
    Set symbol = jobObj.CreateSymbolObject()
    
    symbol.SetId(oooSymbolId)
    
    e3AppObj.PutInfo 0, "=== ЗАПИСЬ АТРИБУТОВ В СИМВОЛ OOO" & number & " ==="
    
    ' Атрибуты от QF устройства
    If Len("" & qfTechDesc) > 0 Then
        symbol.SetAttributeValue "ОД V_Inom", qfTechDesc
        e3AppObj.PutInfo 0, "Записано в ОД V_Inom: " & qfTechDesc
    End If
    
    If Len("" & qfCompTitle) > 0 Then
        symbol.SetAttributeValue "ОД V_Type", qfCompTitle
        e3AppObj.PutInfo 0, "Записано в ОД V_Type: " & qfCompTitle
        
        ' Извлечение и запись значения "kA"
        Dim extractedkAValue
        extractedkAValue = ExtractkAValue(qfCompTitle)
        If Len("" & extractedkAValue) > 0 Then
            symbol.SetAttributeValue "ОД V_Icu", extractedkAValue
            e3AppObj.PutInfo 0, "Записано в ОД V_Icu: " & extractedkAValue
        End If
    End If
    
    If Len("" & qfCompSupplier) > 0 Then
        symbol.SetAttributeValue "ОД V_Proizv", qfCompSupplier
        e3AppObj.PutInfo 0, "Записано в ОД V_Proizv: " & qfCompSupplier
    End If
    
    If Len("" & qfContactCompName) > 0 Then
        symbol.SetAttributeValue "ОД V_Dop ystr", qfContactCompName
        e3AppObj.PutInfo 0, "Записано в ОД V_Dop ystr: " & qfContactCompName
    End If
    
    ' Атрибуты от KM устройства
    If Len("" & kmCompTitle) > 0 Then
        symbol.SetAttributeValue "ОД K_Type", kmCompTitle
        e3AppObj.PutInfo 0, "Записано в ОД K_Type: " & kmCompTitle
    End If
    
    If Len("" & kmCompSupplier) > 0 Then
        symbol.SetAttributeValue "ОД K_Proizv", kmCompSupplier
        e3AppObj.PutInfo 0, "Записано в ОД K_Proizv: " & kmCompSupplier
    End If
    
    If Len("" & kmCompCurrent) > 0 Then
        symbol.SetAttributeValue "ОД K_Inom", kmCompCurrent
        e3AppObj.PutInfo 0, "Записано в ОД K_Inom: " & kmCompCurrent
    End If
    
    ' Атрибуты от KL устройства
    If Len("" & klCompTitle) > 0 Then
        symbol.SetAttributeValue "ОД D_Proizv1", klCompTitle
        e3AppObj.PutInfo 0, "Записано в ОД D_Proizv1: " & klCompTitle
    End If

    ' Атрибуты от V устройства (НОВАЯ ФУНКЦИОНАЛЬНОСТЬ)
    If Len("" & vCompName) > 0 Then
        symbol.SetAttributeValue "ОД D_Proizv2", vCompName
        e3AppObj.PutInfo 0, "Записано в ОД D_Proizv2: " & vCompName
    End If
    
    Set symbol = Nothing
End Sub

' === Основная процедура === Обработка всех символов OOO и устройств
Sub ProcessAllOOOSymbolsAndDevices()
    Dim e3App, job ' Создаем объекты один раз
    Dim oooSymbols, qfDevices, kmDevices, klDevices, vDevices ' Добавлен vDevices
    Dim oooNumber, oooSymbolId
    Dim qfTechDesc, qfCompTitle, qfCompSupplier, qfContactCompName
    Dim kmCompTitle, kmCompSupplier, kmCompCurrent
    Dim klCompTitle
    Dim vCompName ' Новая переменная для имени компонента V устройства
    Dim qfAutomatDeviceId, qfContactDeviceId, kmContactorDeviceId, klDeviceId, vDeviceId
    Dim deviceId, deviceName, componentName
    Dim key
    Dim oooEInomValue ' Переменная для значения ОД E_Inom
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject() ' Создаем JobObject один раз
    
    Set oooSymbols = CreateObject("Scripting.Dictionary")
    Set qfDevices = CreateObject("Scripting.Dictionary")
    Set kmDevices = CreateObject("Scripting.Dictionary")
    Set klDevices = CreateObject("Scripting.Dictionary")
    Set vDevices = CreateObject("Scripting.Dictionary") ' Инициализация нового словаря
    
    e3App.PutInfo 0, "=== СТАРТ ОБРАБОТКИ ВСЕХ OOO СИМВОЛОВ И УСТРОЙСТВ ==="
    
    ' Находим все символы OOO и устройства, передавая e3App и job
    Call FindAllOOOSymbols(e3App, job, oooSymbols)
    Call FindAllDevices(e3App, job, qfDevices, kmDevices, klDevices, vDevices) ' Передаем vDevices
    
    ' Обрабатываем каждый символ OOO
    For Each oooNumber In oooSymbols.Keys
        oooSymbolId = oooSymbols.Item(oooNumber)
        
        e3App.PutInfo 0, "--- ОБРАБОТКА OOO" & oooNumber & " ---"
        
        ' Получаем значение ОД E_Inom из текущего символа OOO
        oooEInomValue = GetOOOAttributeEInom(e3App, job, oooSymbolId) ' Передаем e3App и job
        e3App.PutInfo 0, "Получено значение ОД E_Inom из OOO" & oooNumber & ": " & oooEInomValue
        
        ' Сброс переменных для текущего номера
        qfTechDesc = ""
        qfCompTitle = ""
        qfCompSupplier = ""
        qfContactCompName = ""
        kmCompTitle = ""
        kmCompSupplier = ""
        kmCompCurrent = ""
        klCompTitle = ""
        vCompName = "" ' Сброс для V устройства
        
        qfAutomatDeviceId = 0
        qfContactDeviceId = 0
        kmContactorDeviceId = 0
        klDeviceId = 0
        vDeviceId = 0 ' Сброс для V устройства
        
        ' Поиск соответствующих QF устройств
        If qfDevices.Exists(oooNumber) Then
            For Each key In qfDevices.Item(oooNumber).Keys
                deviceId = CLng(key)
                deviceName = qfDevices.Item(oooNumber).Item(key)
                
                ' Получаем атрибуты, включая имя компонента, для текущего QF устройства
                ' techDesc, compTitle, compSupplier будут перезаписаны, если несколько QF
                Call GetDeviceAttributes(job, deviceId, qfTechDesc, qfCompTitle, qfCompSupplier, Empty, componentName) ' Передаем job
                
                ' НОВАЯ ФУНКЦИОНАЛЬНОСТЬ: Запись ОД E_Inom в Тех. описание 3 для всех QF устройств с данным номером,
                ' ЕСЛИ В ИМЕНИ КОМПОНЕНТА НЕТ "ДИФ"
                If Len("" & oooEInomValue) > 0 Then
                    ' Вызываем новую процедуру, которая содержит проверку на "ДИФ"
                    Call WriteToQFDeviceTechDesc3(e3App, job, deviceId, oooEInomValue) ' Передаем e3App и job
                End If

                If InStr(1, LCase(componentName), "автомат") > 0 Then
                    qfAutomatDeviceId = deviceId
                    e3App.PutInfo 0, "Найден QF Автомат" & oooNumber & ": " & deviceName
                ElseIf InStr(1, LCase(componentName), "контакт") > 0 Then
                    qfContactDeviceId = deviceId
                    qfContactCompName = qfCompTitle ' Используем qfCompTitle для "ОД V_Dop ystr"
                    e3App.PutInfo 0, "Найден QF Контакт" & oooNumber & ": " & deviceName
                End If
                
            Next
        Else
            e3App.PutInfo 0, "QF" & oooNumber & " не найден"
        End If
        
        ' Поиск соответствующих KM устройств
        If kmDevices.Exists(oooNumber) Then
            For Each key In kmDevices.Item(oooNumber).Keys
                deviceId = CLng(key)
                deviceName = kmDevices.Item(oooNumber).Item(key)
                
                ' techDesc не нужен для KM, поэтому передаем Empty
                Call GetDeviceAttributes(job, deviceId, Empty, kmCompTitle, kmCompSupplier, kmCompCurrent, componentName) ' Передаем job
                
                If InStr(1, LCase(componentName), "контактор") > 0 Then
                    kmContactorDeviceId = deviceId
                    e3App.PutInfo 0, "Найден KM Контактор" & oooNumber & ": " & deviceName
                    Exit For ' Берем только первый найденный контактор KM
                End If
            Next
        Else
            e3App.PutInfo 0, "KM" & oooNumber & " не найден"
        End If
        
        ' Поиск соответствующих KL устройств
        If klDevices.Exists(oooNumber) Then
            For Each key In klDevices.Item(oooNumber).Keys
                deviceId = CLng(key)
                deviceName = klDevices.Item(oooNumber).Item(key)
                
                ' techDesc, compSupplier, compCurrent не нужны для KL, поэтому передаем Empty
                Call GetDeviceAttributes(job, deviceId, Empty, klCompTitle, Empty, Empty, componentName) ' Передаем job
                
                klDeviceId = deviceId
                e3App.PutInfo 0, "Найден KL" & oooNumber & ": " & deviceName
                Exit For ' Берем первое найденное KL устройство
            Next
        Else
            e3App.PutInfo 0, "KL" & oooNumber & " не найден"
        End If

        ' Поиск соответствующих V устройств (НОВАЯ ФУНКЦИОНАЛЬНОСТЬ)
        If vDevices.Exists(oooNumber) Then
            For Each key In vDevices.Item(oooNumber).Keys
                deviceId = CLng(key)
                deviceName = vDevices.Item(oooNumber).Item(key)
                
                ' techDesc, compSupplier, compCurrent не нужны для V, поэтому передаем Empty
                Call GetDeviceAttributes(job, deviceId, Empty, Empty, Empty, Empty, vCompName) ' Передаем job, получаем только compName
                
                vDeviceId = deviceId
                e3App.PutInfo 0, "Найден V" & oooNumber & ": " & deviceName
                Exit For ' Берем первое найденное V устройство
            Next
        Else
            e3App.PutInfo 0, "V" & oooNumber & " не найден"
        End If
        
        ' Записываем атрибуты в символ OOO (передаем e3App, job и новый vCompName)
        Call WriteAttributesToOOOSymbol(e3App, job, oooSymbolId, oooNumber, _
                                        qfTechDesc, qfCompTitle, qfCompSupplier, qfContactCompName, _
                                        kmCompTitle, kmCompSupplier, kmCompCurrent, _
                                        klCompTitle, vCompName)
    Next
    
    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ ОБРАБОТКИ ВСЕХ OOO СИМВОЛОВ ==="
    
    Set oooSymbols = Nothing
    Set qfDevices = Nothing
    Set kmDevices = Nothing
    Set klDevices = Nothing
    Set vDevices = Nothing ' Освобождаем новый словарь
    Set job = Nothing ' Освобождаем JobObject
    Set e3App = Nothing ' Освобождаем E3.series Application Object
End Sub

' === Основной запуск ===
Dim e3App_main ' Используем другое имя, чтобы избежать конфликтов с параметрами функций

Set e3App_main = CreateObject("CT.Application")

e3App_main.PutInfo 0, "=== СТАРТ СКРИПТА E3.SERIES OOO PROCESSOR ==="
Call ProcessAllOOOSymbolsAndDevices()
e3App_main.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="

Set e3App_main = Nothing