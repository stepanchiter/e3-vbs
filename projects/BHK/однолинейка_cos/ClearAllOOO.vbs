Option Explicit

' === Функция === Извлечение номера из имени символа
Function ExtractNumber(ByVal itemName)
    Dim re, matches
    Set re = New RegExp
    ' Ищем число в конце строки после символов (например, OOS)
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

' === Процедура === Поиск всех символов OOS в проекте
Sub FindAllOOSSymbols(ByRef OOSSymbols)
    Dim e3App, job, symbol
    Dim symbolIds(), symbolCount
    Dim i, symbolName, symbolNumber
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    
    e3App.PutInfo 0, "=== ПОИСК ВСЕХ СИМВОЛОВ OOS В ПРОЕКТЕ ==="
    
    symbolCount = job.GetSymbolIds(symbolIds)
    If symbolCount = 0 Then
        e3App.PutInfo 0, "В проекте не найдено символов."
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    
    For i = 1 To symbolCount
        symbol.SetId(symbolIds(i))
        symbolName = symbol.GetName()
        
        If LCase(Left(symbolName, 3)) = "OOS" Then
            symbolNumber = ExtractNumber(symbolName)
            If symbolNumber > 0 Then
                OOSSymbols.Add symbolNumber, symbolIds(i)
                e3App.PutInfo 0, "Найден символ OOS: " & symbolName & " (номер: " & symbolNumber & ", ID: " & symbolIds(i) & ")"
            Else
                e3App.PutInfo 0, "Символ OOS найден, но номер не определен: " & symbolName
            End If
        End If
    Next
    
    e3App.PutInfo 0, "Всего найдено символов OOS с номерами: " & OOSSymbols.Count
    
    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === Процедура === Очистка атрибутов символа OOS
Sub ClearOOSSymbolAttributes(ByVal OOSSymbolId, ByVal number)
    Dim e3App, job, symbol
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    
    symbol.SetId(OOSSymbolId)
    
    e3App.PutInfo 0, "=== ОЧИСТКА АТРИБУТОВ СИМВОЛА OOS" & number & " ==="
    
    ' Очистка атрибутов QF устройства
    symbol.SetAttributeValue "ОД V_Inom", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД V_Inom"
    
    symbol.SetAttributeValue "ОД V_Type", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД V_Type"
    
    symbol.SetAttributeValue "ОД V_Icu", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД V_Icu"
    
    symbol.SetAttributeValue "ОД V_Proizv", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД V_Proizv"
    
    symbol.SetAttributeValue "ОД V_Dop ystr", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД V_Dop ystr"
    
    ' Очистка атрибутов KM устройства
    symbol.SetAttributeValue "ОД K_Type", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД K_Type"
    
    symbol.SetAttributeValue "ОД K_Proizv", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД K_Proizv"
    
    symbol.SetAttributeValue "ОД K_Inom", "-"
    e3App.PutInfo 0, "Очищен атрибут ОД K_Inom"
    
    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === Основная процедура === Очистка атрибутов всех символов OOS
Sub ClearAllOOSSymbolsAttributes()
    ' Показываем предупреждающее окно
    Dim msgResult
    msgResult = MsgBox("Очистить данные аппаратов?", vbOKCancel + vbQuestion, "Подтверждение")
    
    ' Если пользователь нажал "Отмена", выходим из скрипта
    If msgResult = vbCancel Then
        Exit Sub
    End If
    
    Dim e3App
    Dim OOSSymbols
    Dim OOSNumber, OOSSymbolId
    
    Set e3App = CreateObject("CT.Application")
    Set OOSSymbols = CreateObject("Scripting.Dictionary")
    
    e3App.PutInfo 0, "=== СТАРТ ОЧИСТКИ АТРИБУТОВ ВСЕХ OOS СИМВОЛОВ ==="
    
    ' Находим все символы OOS
    Call FindAllOOSSymbols(OOSSymbols)
    
    If OOSSymbols.Count = 0 Then
        e3App.PutInfo 0, "Символы OOS не найдены. Очистка не требуется."
        Set OOSSymbols = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    
    ' Очищаем атрибуты каждого символа OOS
    For Each OOSNumber In OOSSymbols.Keys
        OOSSymbolId = OOSSymbols.Item(OOSNumber)
        
        e3App.PutInfo 0, "--- ОЧИСТКА OOS" & OOSNumber & " ---"
        
        Call ClearOOSSymbolAttributes(OOSSymbolId, OOSNumber)
    Next
    
    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ ОЧИСТКИ ВСЕХ OOS СИМВОЛОВ ==="
    e3App.PutInfo 0, "Обработано символов: " & OOSSymbols.Count
    
    Set OOSSymbols = Nothing
    Set e3App = Nothing
End Sub

' === Основной запуск ===
Dim e3App
Set e3App = CreateObject("CT.Application")

e3App.PutInfo 0, "=== СТАРТ СКРИПТА ОЧИСТКИ АТРИБУТОВ OOS СИМВОЛОВ ==="
Call ClearAllOOSSymbolsAttributes()
e3App.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="

Set e3App = Nothing