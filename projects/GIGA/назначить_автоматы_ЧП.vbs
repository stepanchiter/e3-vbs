Option Explicit

' Объявление глобальных словарей
Dim global_oooArticles       ' Для хранения индексов OOO и извлеченных артикулов
Dim global_articleComponentMap ' Для хранения соответствия Артикул -> Имя Компонента Автомата

' === Вспомогательная функция === Нормализация имени компонента (преобразование похожих кириллических символов в латинские/цифровые)
' Эта функция оставлена в коде, но не будет использоваться при заполнении global_articleComponentMap
' по запросу пользователя. Она может быть полезна для других целей.
Function NormalizeComponentName(ByVal compName)
    Dim newCompName : newCompName = compName

    ' Замена похожих кириллических символов на латинские/цифровые аналоги
    ' Это поможет избежать ошибок из-за разницы в кодировке (например, кириллическая Р vs латинская P)
    newCompName = Replace(newCompName, "А", "A") ' Cyrillic A -> Latin A
    newCompName = Replace(newCompName, "В", "B") ' Cyrillic Ve -> Latin B
    newCompName = Replace(newCompName, "Е", "E") ' Cyrillic Ye -> Latin E
    newCompName = Replace(newCompName, "К", "K") ' Cyrillic Ka -> Latin K
    newCompName = Replace(newCompName, "М", "M") ' Cyrillic Em -> Latin M
    newCompName = Replace(newCompName, "Н", "H") ' Cyrillic En -> Latin H
    newCompName = Replace(newCompName, "О", "O") ' Cyrillic O -> Latin O
    newCompName = Replace(newCompName, "Р", "P") ' Cyrillic Er -> Latin P (особенно важно для "3Р" -> "3P")
    newCompName = Replace(newCompName, "С", "C") ' Cyrillic Es -> Latin C
    newCompName = Replace(newCompName, "Т", "T") ' Cyrillic Te -> Latin T
    newCompName = Replace(newCompName, "Х", "X") ' Cyrillic Kha -> Latin X
    newCompName = Replace(newCompName, "У", "Y") ' Cyrillic U -> Latin Y
    
    ' Специальная замена для "ЗР" -> "3P" (если "З" - это опечатка для "3" в комбинации с "Р")
    newCompName = Replace(newCompName, "ЗР", "3P") 
    ' Также проверим просто З на 3
    newCompName = Replace(newCompName, "З", "3")


    NormalizeComponentName = newCompName
End Function


' === Главная процедура === Извлечение артикулов из атрибутов OOO символов
Sub ExtractOOOArticles()
    Dim e3App, job, symbol
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName
    Dim s
    
    ' Переменные для атрибутов
    Dim dProizv3Value, dProizv2Value
    Dim extractedArticle ' Для хранения извлеченного артикула

    ' Для регулярного выражения
    Dim re, matches
    Set re = New RegExp
    re.Pattern = "(ABA|ABC)\d+" 
    re.IgnoreCase = True  
    re.Global = False     

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Извлечение артикулов OOO символов ==="

    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "В проекте нет символов для анализа. Скрипт завершен."
        Call CleanUpObjects(re, matches, symbol, job, e3App) 
        Exit Sub
    End If

    e3App.PutInfo 0, "Найдено " & allSymbolCount & " символов в проекте. Поиск OOO символов со ОД D_Proizv3=1..."
    
    Dim foundOOOCount : foundOOOCount = 0

    For s = 1 To allSymbolCount 
        currentSymbolId = allSymbolIds(s)
        symbol.SetId(currentSymbolId)
        symbolName = symbol.GetName()

        If LCase(Left(symbolName, 3)) = "ooo" Then
            dProizv3Value = Trim(CStr(symbol.GetAttributeValue("ОД D_Proizv3")))

            If dProizv3Value = "1" Then
                foundOOOCount = foundOOOCount + 1
                e3App.PutInfo 0, "  Найден OOO символ: '" & symbolName & "' (ID: " & currentSymbolId & ") с ОД D_Proizv3 = 1."

                dProizv2Value = Trim(CStr(symbol.GetAttributeValue("ОД D_Proizv2")))
                e3App.PutInfo 0, "    Значение ОД D_Proizv2: '" & dProizv2Value & "'"

                Dim oooIndex
                On Error Resume Next 
                oooIndex = CLng(Mid(symbolName, 4))
                If Err.Number <> 0 Then
                    oooIndex = "Неизвестен" 
                    Err.Clear
                End If
                On Error GoTo 0

                If Len(dProizv2Value) > 0 Then
                    Set matches = re.Execute(dProizv2Value)
                    If matches.Count > 0 Then
                        extractedArticle = matches.Item(0).Value 
                        e3App.PutInfo 0, "    Извлеченный артикул: '" & extractedArticle & "'"
                        
                        If Not global_oooArticles.Exists(CStr(oooIndex)) Then
                            global_oooArticles.Add CStr(oooIndex), extractedArticle
                        Else
                            global_oooArticles.Item(CStr(oooIndex)) = extractedArticle 
                        End If
                    Else
                        extractedArticle = "НЕ НАЙДЕНО"
                        e3App.PutInfo 0, "    Артикул не найден по шаблону в ОД D_Proizv2."
                    End If
                Else
                    extractedArticle = "ПУСТО"
                    e3App.PutInfo 0, "    ОД D_Proizv2 пуст, артикул не может быть извлечен."
                End If
                
                e3App.PutInfo 0, "  --- Результат для OOO" & oooIndex & ": Артикул = " & extractedArticle & " ---"
            End If
        End If
    Next 

    If foundOOOCount = 0 Then
        e3App.PutInfo 0, "Не найдено OOO символов со значением атрибута ОД D_Proizv3 равным '1'."
    End If

    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ СКРИПТА ==="

    Call CleanUpObjects(re, matches, symbol, job, e3App)
End Sub

' === Вспомогательная процедура === Очистка объектов
Sub CleanUpObjects(reObj, matchesObj, symbolObj, jobObj, appObj)
    Set reObj = Nothing
    Set matchesObj = Nothing
    Set symbolObj = Nothing
    Set jobObj = Nothing
    Set appObj = Nothing
End Sub

' === ПРОЦЕДУРА FindQFKM (не изменена, оставлена для полноты) ===
Sub FindQFKM(job)
    Dim device, re, matches
    Dim deviceIds, deviceCount
    Dim i, j
    Dim devName, compName
    Dim deviceInfo
    Dim key, arr, k
    Dim e3App_local 

    Set e3App_local = CreateObject("CT.Application") 

    Set device = job.CreateDeviceObject()
    Set re = New RegExp
    re.Pattern = "^-(QF|KM)\d+$" 
    re.IgnoreCase = False
    re.Global = False

    deviceCount = job.GetAllDeviceIds(deviceIds)

    If deviceCount = 0 Then
        e3App_local.PutInfo 0, "Устройства QF/KM не найдены в проекте." 
        Set device = Nothing
        Set re = Nothing
        Set e3App_local = Nothing 
        Exit Sub
    End If

    Set deviceInfo = CreateObject("Scripting.Dictionary")

    e3App_local.PutInfo 0, "=== FindQFKM: Начало поиска устройств -QF с компонентом 'Автомат' и всех -KM устройств ===" 

    For i = 1 To deviceCount
        device.SetId(deviceIds(i))
        devName = device.GetName()

        If re.Test(devName) Then 
            compName = device.GetComponentName()
            
            If LCase(Left(devName, 3)) = "-qf" Then 
                If InStr(1, LCase(compName), "автомат") > 0 Then 
                    If Not deviceInfo.Exists(devName) Then
                        deviceInfo.Add devName, Array() 
                    End If
                    Dim oldArr_qf
                    oldArr_qf = deviceInfo.Item(devName) 
                    Dim currentSize_qf : currentSize_qf = -1
                    On Error Resume Next 
                    currentSize_qf = UBound(oldArr_qf)
                    On Error GoTo 0 
                    Dim newArr_qf
                    If currentSize_qf = -1 Then 
                        ReDim newArr_qf(0) 
                    Else
                        ReDim Preserve newArr_qf(currentSize_qf + 1)
                    End If
                    For j = LBound(oldArr_qf) To currentSize_qf 
                        newArr_qf(j) = oldArr_qf(j)
                    Next
                    newArr_qf(UBound(newArr_qf)) = "ID=" & deviceIds(i) & ", Component=" & compName
                    deviceInfo.Item(devName) = newArr_qf
                Else
                    e3App_local.PutInfo 0, "  Пропущено устройство: " & devName & " (ID: " & deviceIds(i) & ") - компонент '" & compName & "' не содержит 'Автомат'."
                End If
            ElseIf LCase(Left(devName, 3)) = "-km" Then 
                If Not deviceInfo.Exists(devName) Then
                    deviceInfo.Add devName, Array() 
                End If
                Dim oldArr_km
                oldArr_km = deviceInfo.Item(devName) 
                Dim currentSize_km : currentSize_km = -1
                On Error Resume Next 
                currentSize_km = UBound(oldArr_km)
                On Error GoTo 0 
                Dim newArr_km
                If currentSize_km = -1 Then 
                    ReDim newArr_km(0) 
                Else
                    ReDim Preserve newArr_km(currentSize_km + 1)
                End If
                For j = LBound(oldArr_km) To currentSize_km 
                    newArr_km(j) = oldArr_km(j)
                Next
                newArr_km(UBound(newArr_km)) = "ID=" & deviceIds(i) & ", Component=" & compName
                deviceInfo.Item(devName) = newArr_km
            End If 
        End If 
    Next

    e3App_local.PutInfo 0, "=== FindQFKM: найдено " & deviceInfo.Count & " групп устройств ===" 

    If deviceInfo.Count > 0 Then
        For Each key In deviceInfo.Keys
            e3App_local.PutInfo 0, "Устройство: " & key 
            arr = deviceInfo.Item(key)
            For k = LBound(arr) To UBound(arr)
                e3App_local.PutInfo 0, "    " & arr(k) 
            Next
        Next
    Else
        e3App_local.PutInfo 0, "Не найдено устройств, соответствующих шаблону -QF### (с 'Автомат' в компоненте) или -KM###." 
    End If

    Set device = Nothing
    Set re = Nothing
    Set deviceInfo = Nothing
    Set e3App_local = Nothing 
End Sub


' === НОВАЯ ПРОЦЕДУРА === Обновление имени компонента QF на основе артикула OOO
Sub COMM(job, e3App_local)
    Dim device
    Dim oooIndex, extractedArticle
    Dim targetDeviceName, targetDeviceId
    Dim currentDeviceName, currentComponentName
    Dim allDeviceIds(), deviceCount
    Dim i
    Dim newComponentName 
    Dim componentVersion 

    Set device = job.CreateDeviceObject()

    e3App_local.PutInfo 0, "=== COMM: Начало процедуры обновления компонентов ==="

    If global_oooArticles.Count = 0 Then
        e3App_local.PutInfo 0, "COMM: Нет данных об OOO символах с ОД D_Proizv3=1 для обработки."
        Set device = Nothing 
        Exit Sub
    End If

    For Each oooIndex In global_oooArticles.Keys 
        extractedArticle = global_oooArticles.Item(oooIndex) 
        
        targetDeviceName = "-QF" & oooIndex 
        e3App_local.PutInfo 0, "  Поиск устройства '" & targetDeviceName & "' для обновления (OOO" & oooIndex & ", Артикул: " & extractedArticle & ")..."

        targetDeviceId = 0 
        deviceCount = job.GetAllDeviceIds(allDeviceIds) 
        
        For i = 1 To deviceCount
            device.SetId(allDeviceIds(i))
            currentDeviceName = device.GetName()
            currentComponentName = device.GetComponentName()

            If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                If InStr(1, LCase(currentComponentName), "автомат") > 0 Then
                    targetDeviceId = allDeviceIds(i)
                    e3App_local.PutInfo 0, "    Найдено устройство '" & targetDeviceName & "' (ID: " & targetDeviceId & ") с компонентом 'Автомат'."
                    Exit For 
                Else
                    e3App_local.PutInfo 0, "    Устройство '" & targetDeviceName & "' найдено, но его компонент ('" & currentComponentName & "') не содержит 'Автомат'. Пропускаем."
                End If
            End If
        Next

        If targetDeviceId > 0 Then
            device.SetId(targetDeviceId)
            
            If global_articleComponentMap.Exists(extractedArticle) Then
                ' Получаем имя компонента из словаря (без нормализации здесь, так как она уже сделана при заполнении)
                newComponentName = global_articleComponentMap.Item(extractedArticle) 
                componentVersion = "1" 
                
                On Error Resume Next 
                device.SetComponentName newComponentName, componentVersion 
                If Err.Number = 0 Then
                    e3App_local.PutInfo 0, "    УСПЕШНО: Компонент устройства '" & targetDeviceName & "' обновлен на: '" & newComponentName & "' (Версия: '" & componentVersion & "')."
                Else
                    e3App_local.PutInfo 0, "    ОШИБКА при обновлении компонента для '" & targetDeviceName & "': " & Err.Description
                    Err.Clear 
                End If
                On Error GoTo 0 
            Else
                e3App_local.PutInfo 0, "    ПРЕДУПРЕЖДЕНИЕ: Для артикула '" & extractedArticle & "' (OOO" & oooIndex & ") не найдено соответствия в таблице компонентов. Компонент не обновлен."
            End If
        Else
            e3App_local.PutInfo 0, "    ПРЕДУПРЕЖДЕНИЕ: Устройство '" & targetDeviceName & "' (с компонентом 'Автомат') не найдено в проекте для OOO" & oooIndex & ". Обновление пропущено."
        End If
    Next

    e3App_local.PutInfo 0, "=== COMM: Завершение процедуры обновления компонентов ==="

    Set device = Nothing 
End Sub

' === Основной запуск ===
Dim global_e3App, global_job
Set global_e3App = CreateObject("CT.Application")
Set global_job = global_e3App.CreateJobObject()

Set global_oooArticles = CreateObject("Scripting.Dictionary") 
Set global_articleComponentMap = CreateObject("Scripting.Dictionary") 

' Заполнение global_articleComponentMap данными из предоставленной текстовой таблицы,
' включая image_6c9280.png, image_917c55.png, image_917855.png и новые данные из последнего запроса.
' БЕЗ применения NormalizeComponentName к значениям (именам компонентов), по запросу пользователя.
global_articleComponentMap.Add "ABA00002", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABA00110", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABA00005", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABA00104", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABA00003", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABA00006", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABA00105", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABA00011", "Автомат_3Р_50A_13182DEK"
global_articleComponentMap.Add "ABA00111", "Автомат_3Р_50A_13182DEK"
global_articleComponentMap.Add "ABA00012", "Автомат_3Р_50A_13182DEK"
global_articleComponentMap.Add "ABA00112", "Автомат_3Р_50A_13182DEK"
global_articleComponentMap.Add "ABA00013", "Автомат_3Р_63A_13183DEK"
global_articleComponentMap.Add "ABA00113", "Автомат_3Р_63A_13183DEK"
global_articleComponentMap.Add "ABA00004", "Автомат_3Р_16A_13177DEK"
global_articleComponentMap.Add "ABA00007", "Автомат_3Р_16A_13177DEK"
global_articleComponentMap.Add "ABA00106", "Автомат_3Р_16A_13177DEK"
global_articleComponentMap.Add "ABA00014", "Автомат_3Р_80A_13008DEK"
global_articleComponentMap.Add "ABA00114", "Автомат_3Р_80A_13008DEK"
global_articleComponentMap.Add "ABA00107", "Автомат_3Р_20A_13178DEK"
global_articleComponentMap.Add "ABA00008", "Автомат_3Р_20A_13178DEK"
global_articleComponentMap.Add "ABA00108", "Автомат_3Р_20A_13178DEK"
global_articleComponentMap.Add "ABA00009", "Автомат_3Р_25A_13179DEK"
global_articleComponentMap.Add "ABA00109", "Автомат_3Р_25A_13179DEK"
global_articleComponentMap.Add "ABA00010", "Автомат_3Р_25A_13179DEK"
global_articleComponentMap.Add "ABC00023", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABC00024", "Автомат_3Р_10A_13176DEK"
global_articleComponentMap.Add "ABC00029", "Автомат_3Р_40A_13181DEK"
global_articleComponentMap.Add "ABC00030", "Автомат_3Р_50A_13182DEK"
global_articleComponentMap.Add "ABC00031", "Автомат_3Р_63A_13183DEK"
global_articleComponentMap.Add "ABC00025", "Автомат_3Р_16A_13177DEK"
global_articleComponentMap.Add "ABC00032", "Автомат_3Р_80A_13008DEK"
global_articleComponentMap.Add "ABC00033", "Автомат_3Р_100A_13009DEK"
global_articleComponentMap.Add "ABC00034", "Автомат_3Р_125A_13027DEK"
global_articleComponentMap.Add "ABC00035", "Автомат_3Р_160A_22752DEK"
global_articleComponentMap.Add "ABC00027", "Автомат_3Р_25A_13179DEK"
global_articleComponentMap.Add "ABC00066", "Автомат_3Р_200A_22754DEK"
global_articleComponentMap.Add "ABC00028", "Автомат_3Р_32A_13180DEK"
global_articleComponentMap.Add "ABC00037", "Автомат_3Р_200A_22754DEK"
global_articleComponentMap.Add "ABC00067", "Автомат_3Р_200A_22754DEK"
global_articleComponentMap.Add "ABC00038", "Автомат_3Р_250A_22756DEK"
global_articleComponentMap.Add "ABC00068", "Автомат_3Р_250A_22756DEK"
global_articleComponentMap.Add "ABC00026", "Автомат_3Р_20A_13178DEK"
global_articleComponentMap.Add "ABA00001", "Автомат_3Р_40A_13181DEK"
global_articleComponentMap.Add "ABC00036", "Автомат_3Р_160A_22752DEK"
global_articleComponentMap.Add "ABC00123", "Автомат_3Р_10A_13176DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00124", "Автомат_3Р_10A_13176DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00129", "Автомат_3Р_40A_13181DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00130", "Автомат_3Р_50A_13182DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00131", "Автомат_3Р_63A_13183DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00125", "Автомат_3Р_16A_13177DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00132", "Автомат_3Р_80A_13008DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00133", "Автомат_3Р_100A_13009DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00134", "Автомат_3Р_125A_13027DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00135", "Автомат_3Р_160A_22752DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00127", "Автомат_3Р_25A_13179DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00166", "Автомат_3Р_200A_22754DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00128", "Автомат_3Р_32A_13180DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00137", "Автомат_3Р_200A_22754DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00167", "Автомат_3Р_200A_22754DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00138", "Автомат_3Р_250A_22756DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00168", "Автомат_3Р_250A_22756DEK" ' <-- ДОБАВЛЕНО
global_articleComponentMap.Add "ABC00126", "Автомат_3Р_20A_13178DEK" ' <-- ДОБАВЛЕНО


' Вызываем основные процедуры
Call ExtractOOOArticles() 
Call FindQFKM(global_job) 
Call COMM(global_job, global_e3App) 

' Финальная очистка глобальных объектов
Set global_job = Nothing
Set global_e3App = Nothing
Set global_oooArticles = Nothing 
Set global_articleComponentMap = Nothing 
