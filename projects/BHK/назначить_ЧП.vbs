'*******************************************************************************
' Название скрипта: E3_UZ_ComponentUpdater
' Автор: E3.series VBScript Assistant
' Дата: 01.07.2025
' Описание: Скрипт для автоматического обновления имен компонентов для устройств -UZ
'          на основе артикулов, извлеченных из OOO символов, и новой таблицы соответствия.
'          Не проверяет имя компонента на "автомат".
'*******************************************************************************

Option Explicit

'*******************************************************************************
' Глобальные переменные
'*******************************************************************************
Dim global_e3App             ' Объект приложения E3.series
Dim global_job               ' Объект текущего проекта E3.series
Dim global_oooArticles       ' Словарь для хранения индексов OOO и извлеченных артикулов (OOOIndex -> ExtractedArticle)
Dim global_articleComponentMap ' Словарь для хранения соответствия Артикул -> Новое Имя Компонента UZ

'*******************************************************************************
' Процедура ExtractOOOArticles()
' Ищет символы OOO, извлекает из них артикулы и сохраняет в global_oooArticles.
'*******************************************************************************
Sub ExtractOOOArticles()
    Dim symbolIds, symbolCount, i
    Dim symbol, symbolName, attributeValue, extractedArticle, oooIndex
    Dim regEx, matches

    Set symbol = global_job.CreateSymbolObject()
    Set regEx = New RegExp
    regEx.IgnoreCase = True ' Для поиска OOO, ABA, ABC без учета регистра
    regEx.Global = True

    global_e3App.PutInfo 0, "Начинаем поиск OOO символов и извлечение артикулов..."

    ' ИСПРАВЛЕНИЕ: Используем GetSymbolIds вместо GetAllSymbolIds
    symbolCount = global_job.GetSymbolIds(symbolIds) 

    If symbolCount > 0 Then
        For i = 1 To symbolCount
            symbol.SetId(symbolIds(i))
            symbolName = symbol.GetName()

            ' Проверяем, является ли символ OOO символом (например, OOO1, OOO10)
            If LCase(Left(symbolName, 3)) = "ooo" Then
                ' Проверяем атрибут "ОД D_Proizv3" (Manufacturer Specific Data 3)
                attributeValue = symbol.GetAttributeValue("ОД D_Proizv3")

                If attributeValue = "1" Then
                    ' Извлекаем артикул из атрибута "ОД D_Proizv2" (Manufacturer Specific Data 2)
                    attributeValue = symbol.GetAttributeValue("ОД D_Proizv2")

                    regEx.Pattern = "(ABA|ABC)\d+" ' Паттерн для артикулов
                    Set matches = regEx.Execute(attributeValue)

                    If matches.Count > 0 Then
                        extractedArticle = matches(0).Value
                        ' Извлекаем числовой индекс из имени символа OOO (например, из OOO123 -> 123)
                        regEx.Pattern = "\d+"
                        Set matches = regEx.Execute(symbolName)
                        If matches.Count > 0 Then
                            oooIndex = CInt(matches(0).Value)
                            If Not global_oooArticles.Exists(oooIndex) Then
                                global_oooArticles.Add oooIndex, extractedArticle
                                global_e3App.PutInfo 0, "Найден OOO символ: " & symbolName & _
                                    ", Атрибут 'ОД D_Proizv3' = '" & symbol.GetAttributeValue("ОД D_Proizv3") & _
                                    "', Извлечен артикул из 'ОД D_Proizv2': '" & extractedArticle & "'"
                            Else
                                global_e3App.PutInfo 1, "Предупреждение: OOO индекс '" & oooIndex & _
                                    "' уже существует в словаре. Пропускаем дубликат: " & symbolName
                            End If
                        Else
                            global_e3App.PutInfo 1, "Предупреждение: Не удалось извлечь числовой индекс из имени OOO символа: " & symbolName
                        End If
                    Else
                        global_e3App.PutInfo 1, "Предупреждение: В атрибуте 'ОД D_Proizv2' символа '" & symbolName & _
                            "' не найден артикул по шаблону (ABA|ABC)######."
                    End If
                Else
                    global_e3App.PutInfo 0, "Пропускаем символ '" & symbolName & _
                        "', так как атрибут 'ОД D_Proizv3' не равен '1'."
                End If
            End If
        Next
    Else
        global_e3App.PutInfo 1, "В проекте не найдено символов."
    End If

    global_e3App.PutInfo 0, "Поиск OOO символов завершен. Найдено артикулов: " & global_oooArticles.Count
    Call CleanUpObjects(regEx, matches, symbol, Nothing, Nothing) ' matches и symbol очищаются здесь
End Sub

'*******************************************************************************
' Процедура CleanUpObjects()
' Освобождает COM-объекты.
'*******************************************************************************
Sub CleanUpObjects(ByRef reObj, ByRef matchesObj, ByRef symbolObj, ByRef jobObj, ByRef appObj)
    If Not reObj Is Nothing Then Set reObj = Nothing
    If Not matchesObj Is Nothing Then Set matchesObj = Nothing
    If Not symbolObj Is Nothing Then Set symbolObj = Nothing
    If Not jobObj Is Nothing Then Set jobObj = Nothing
    If Not appObj Is Nothing Then Set appObj = Nothing
End Sub

'*******************************************************************************
' Процедура FindUZDevices(job)
' Ищет устройства -UZ### и выводит о них информацию.
' (Эта процедура служит для диагностики/логирования и не передает данные для изменения.)
'*******************************************************************************
Sub FindUZDevices(job)
    Dim deviceIds, deviceCount, i
    Dim device, devName
    Dim regEx, deviceInfo ' deviceInfo - локальный словарь для этой процедуры

    Set device = job.CreateDeviceObject()
    Set regEx = New RegExp
    regEx.Pattern = "^-UZ\d+$" ' Паттерн для поиска -UZ###
    regEx.IgnoreCase = True
    Set deviceInfo = CreateObject("Scripting.Dictionary")

    global_e3App.PutInfo 0, "Начинаем поиск устройств -UZ..."

    deviceCount = job.GetAllDeviceIds(deviceIds)

    If deviceCount > 0 Then
        For i = 1 To deviceCount
            device.SetId(deviceIds(i))
            devName = device.GetName()

            If regEx.Test(devName) Then
                ' Для -UZ устройств не требуется проверка имени компонента на "автомат"
                If Not deviceInfo.Exists(devName) Then
                    deviceInfo.Add devName, device.GetComponentName()
                    global_e3App.PutInfo 0, "Найден девайс -UZ: " & devName & _
                        ", Текущий компонент: '" & device.GetComponentName() & "'"
                End If
            End If
        Next
    Else
        global_e3App.PutInfo 1, "В проекте не найдено устройств."
    End If

    If deviceInfo.Count = 0 Then
        global_e3App.PutInfo 1, "Внимание: Устройства -UZ, соответствующие шаблону, не найдены."
    End If

    global_e3App.PutInfo 0, "Поиск устройств -UZ завершен."
    Set deviceInfo = Nothing
    Call CleanUpObjects(regEx, Nothing, device, Nothing, Nothing)
End Sub

'*******************************************************************************
' Процедура UpdateUZComponents(job, e3App_local)
' Обновляет имя компонента для устройств -UZ на основе таблицы соответствия.
'*******************************************************************************
Sub UpdateUZComponents(job, e3App_local)
    Dim deviceIds, deviceCount, i, oooIndex
    Dim device, devName, currentComponentName, newComponentName, extractedArticle
    Dim regEx, matches
    Dim targetUZName
    Dim componentVersion ' Объявляем переменную для версии компонента

    ' Устанавливаем версию компонента, как вы указали
    componentVersion = "1" 

    Set device = job.CreateDeviceObject()
    Set regEx = New RegExp
    regEx.IgnoreCase = True

    e3App_local.PutInfo 0, "Начинаем обновление компонентов для устройств -UZ..."

    ' Итерируем по всем извлеченным артикулам из OOO символов
    For Each oooIndex In global_oooArticles.Keys()
        extractedArticle = global_oooArticles.Item(oooIndex)
        targetUZName = "-UZ" & oooIndex ' Формируем имя целевого -UZ устройства

        e3App_local.PutInfo 0, "Обрабатываем OOO-артикул: '" & extractedArticle & "' для потенциального -UZ устройства: '" & targetUZName & "'"

        ' Ищем целевое -UZ устройство в проекте
        deviceCount = job.GetAllDeviceIds(deviceIds)
        Dim foundTargetDevice : foundTargetDevice = False

        If deviceCount > 0 Then
            For i = 1 To deviceCount
                device.SetId(deviceIds(i))
                devName = device.GetName()

                ' Проверяем, соответствует ли имя устройства целевому -UZ
                If LCase(devName) = LCase(targetUZName) Then
                    currentComponentName = device.GetComponentName()
                    e3App_local.PutInfo 0, "Найден девайс: " & devName & ", текущий компонент: '" & currentComponentName & "'"

                    ' Ищем новое имя компонента в global_articleComponentMap
                    If global_articleComponentMap.Exists(extractedArticle) Then
                        newComponentName = global_articleComponentMap.Item(extractedArticle)

                        If LCase(currentComponentName) <> LCase(newComponentName) Then
                            On Error Resume Next ' Включаем обработку ошибок для SetComponentName
                            ' ИСПРАВЛЕНИЕ: Передаем имя компонента и версию
                            device.SetComponentName newComponentName, componentVersion
                            If Err.Number = 0 Then
                                e3App_local.PutInfo 0, "УСПЕШНО: Обновлен компонент девайса '" & devName & _
                                    "' с '" & currentComponentName & "' на '" & newComponentName & "' (версия: " & componentVersion & ")"
                            Else
                                e3App_local.PutInfo 2, "ОШИБКА: Не удалось обновить компонент девайса '" & devName & _
                                    "' на '" & newComponentName & "'. Ошибка: " & Err.Description & _
                                    " (Код: " & Err.Number & ", Источник: " & Err.Source & ")" ' Добавил Err.Source для лучшей диагностики
                            End If
                            On Error GoTo 0 ' Выключаем обработку ошибок
                        Else
                            e3App_local.PutInfo 0, "Девайс '" & devName & "' уже имеет требуемый компонент: '" & newComponentName & "'. Пропускаем."
                        End If
                    Else
                        e3App_local.PutInfo 1, "ПРЕДУПРЕЖДЕНИЕ: Артикул '" & extractedArticle & _
                            "' для девайса '" & devName & "' не найден в таблице соответствия global_articleComponentMap. Компонент не обновлен."
                    End If
                    foundTargetDevice = True
                    Exit For ' Девайс найден и обработан, переходим к следующему OOO-артикулу
                End If
            Next
        End If

        If Not foundTargetDevice Then
            e3App_local.PutInfo 1, "ПРЕДУПРЕЖДЕНИЕ: Целевой девайс '" & targetUZName & "' не найден в проекте для артикула '" & extractedArticle & "'. Пропускаем."
        End If

    Next

    e3App_local.PutInfo 0, "Обновление компонентов для устройств -UZ завершено."
    Call CleanUpObjects(regEx, Nothing, device, Nothing, Nothing)
End Sub

'*******************************************************************************
' Основной блок выполнения скрипта
'*******************************************************************************
Sub Main()

    ' 1. Инициализация объектов E3.series
    Set global_e3App = CreateObject("CT.Application")
    If global_e3App Is Nothing Then
        MsgBox "Не удалось подключиться к E3.series. Убедитесь, что E3.series запущен.", vbCritical, "Ошибка E3.series"
        Exit Sub
    End If
    Set global_job = global_e3App.CreateJobObject()
    If global_job Is Nothing Then
        MsgBox "Не удалось получить доступ к текущему проекту E3.series.", vbCritical, "Ошибка E3.series"
        Call CleanUpObjects(Nothing, Nothing, Nothing, Nothing, global_e3App)
        Exit Sub
    End If

    ' Очистка окна сообщений E3.series в начале выполнения
    global_e3App.PutMessageEx 0, "Запуск скрипта E3_UZ_ComponentUpdater...", 0, 0, 0, 249 ' Цвет BLUE

    ' 2. Инициализация глобальных словарей
    Set global_oooArticles = CreateObject("Scripting.Dictionary")
    Set global_articleComponentMap = CreateObject("Scripting.Dictionary")

    ' 3. Заполнение global_articleComponentMap
    ' ЭТОТ РАЗДЕЛ БУДЕТ ЗАПОЛНЕН ВАШИМИ ДАННЫМИ ПОЗЖЕ.
    ' Пример: global_articleComponentMap.Add "ВАШ_АРТИКУЛ", "ВАШ_НОВЫЙ_КОМПОНЕНТ"
    global_articleComponentMap.Add "ABA00005", "VF51_0.75"
    global_articleComponentMap.Add "ABA00006", "VF51_1.5"
    global_articleComponentMap.Add "ABA00011", "VF51_11.0"
    global_articleComponentMap.Add "ABA00012", "VF51_15.0"
    global_articleComponentMap.Add "ABA00013", "VF51_18.5"
    global_articleComponentMap.Add "ABA00007", "VF51_2.2"
    global_articleComponentMap.Add "ABA00014", "VF51_22.0"
    global_articleComponentMap.Add "ABA00107", "VF51_3.0"
    global_articleComponentMap.Add "ABA00008", "VF51_4.0"
    global_articleComponentMap.Add "ABA00009", "VF51_5.5"
    global_articleComponentMap.Add "ABA00010", "VF51_7.5"
    global_articleComponentMap.Add "ABC00123", "VF101_0.75"
    global_articleComponentMap.Add "ABC00124", "VF101_1.5"
    global_articleComponentMap.Add "ABC00129", "VF101_11"
    global_articleComponentMap.Add "ABC00130", "VF101_15"
    global_articleComponentMap.Add "ABC00131", "VF101_18.5"
    global_articleComponentMap.Add "ABC00125", "VF101_2.2"
    global_articleComponentMap.Add "ABC00132", "VF101_22"
    global_articleComponentMap.Add "ABC00133", "VF101_30"
    global_articleComponentMap.Add "ABC00134", "VF101_37"
    global_articleComponentMap.Add "ABC00135", "VF101_45"
    global_articleComponentMap.Add "ABC00127", "VF101_5.5"
    global_articleComponentMap.Add "ABC00166", "VF101_55"
    global_articleComponentMap.Add "ABC00128", "VF101_7.5"
    global_articleComponentMap.Add "ABC00137", "VF101_75"
    global_articleComponentMap.Add "ABC00138", "VF101_90"
    global_articleComponentMap.Add "ABC00126", "VF101_4"

    ' 4. Выполнение основных процедур
    Call ExtractOOOArticles()        ' Извлекаем артикулы из OOO символов
    Call FindUZDevices(global_job)   ' Ищем устройства -UZ (для логирования/диагностики)
    Call UpdateUZComponents(global_job, global_e3App) ' Обновляем компоненты -UZ устройств

    global_e3App.PutInfo 0, "Скрипт E3_UZ_ComponentUpdater завершен."

    ' 5. Очистка глобальных объектов
    Call CleanUpObjects(Nothing, Nothing, Nothing, global_job, global_e3App)
    Set global_oooArticles = Nothing
    Set global_articleComponentMap = Nothing

    Exit Sub ' Выход из Sub Main при успешном выполнении

ErrorHandler:
    global_e3App.PutInfo 2, "КРИТИЧЕСКАЯ ОШИБКА в скрипте: " & Err.Description & " (Код: " & Err.Number & ")"
    MsgBox "Произошла критическая ошибка в скрипте. См. окно сообщений E3.series для подробностей.", vbCritical, "Ошибка скрипта"
    Call CleanUpObjects(Nothing, Nothing, Nothing, global_job, global_e3App)
    Set global_oooArticles = Nothing
    Set global_articleComponentMap = Nothing
End Sub

' Запуск основной процедуры
Call Main()