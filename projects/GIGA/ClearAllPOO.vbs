Option Explicit

' === Функция === Извлечение номера из имени символа
Function ExtractNumber(ByVal itemName)
    Dim re, matches
    Set re = New RegExp
    ' Ищем число в конце строки после символов (например, POO1)
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

' === Процедура === Поиск всех символов POO в проекте
Sub FindAllPOOSymbols(ByRef POOSymbols)
    Dim e3App, job, symbol
    Dim symbolIds(), symbolCount
    Dim i, symbolName, symbolNumber
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    
    e3App.PutInfo 0, "=== ПОИСК ВСЕХ СИМВОЛОВ POO В ПРОЕКТЕ ==="
    
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
        
        If LCase(Left(symbolName, 3)) = "POO" Then
            symbolNumber = ExtractNumber(symbolName)
            If symbolNumber > 0 Then
                POOSymbols.Add symbolNumber, symbolIds(i)
                e3App.PutInfo 0, "Найден символ POO: " & symbolName & " (номер: " & symbolNumber & ", ID: " & symbolIds(i) & ")"
            Else
                e3App.PutInfo 0, "Символ POO найден, но номер не определен: " & symbolName
            End If
        End If
    Next
    
    e3App.PutInfo 0, "Всего найдено символов POO с номерами: " & POOSymbols.Count
    
    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === Процедура === Очистка атрибутов символа POO
Sub ClearPOOSymbolAttributes(ByVal POOSymbolId, ByVal number)
    Dim e3App, job, symbol
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    
    symbol.SetId(POOSymbolId)
    
    e3App.PutInfo 0, "=== ОЧИСТКА АТРИБУТОВ СИМВОЛА POO" & number & " ==="
    
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

' === Основная процедура === Очистка атрибутов всех символов POO
Sub ClearAllPOOSymbolsAttributes()
    ' Показываем предупреждающее окно
    Dim msgResult
    msgResult = MsgBox("Очистить данные аппаратов?", vbOKCancel + vbQuestion, "Подтверждение")
    
    ' Если пользователь нажал "Отмена", выходим из скрипта
    If msgResult = vbCancel Then
        Exit Sub
    End If
    
    Dim e3App
    Dim POOSymbols
    Dim POONumber, POOSymbolId
    
    Set e3App = CreateObject("CT.Application")
    Set POOSymbols = CreateObject("Scripting.Dictionary")
    
    e3App.PutInfo 0, "=== СТАРТ ОЧИСТКИ АТРИБУТОВ ВСЕХ POO СИМВОЛОВ ==="
    
    ' Находим все символы POO
    Call FindAllPOOSymbols(POOSymbols)
    
    If POOSymbols.Count = 0 Then
        e3App.PutInfo 0, "Символы POO не найдены. Очистка не требуется."
        Set POOSymbols = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    
    ' Очищаем атрибуты каждого символа POO
    For Each POONumber In POOSymbols.Keys
        POOSymbolId = POOSymbols.Item(POONumber)
        
        e3App.PutInfo 0, "--- ОЧИСТКА POO" & POONumber & " ---"
        
        Call ClearPOOSymbolAttributes(POOSymbolId, POONumber)
    Next
    
    e3App.PutInfo 0, "=== ЗАВЕРШЕНИЕ ОЧИСТКИ ВСЕХ POO СИМВОЛОВ ==="
    e3App.PutInfo 0, "Обработано символов: " & POOSymbols.Count
    
    Set POOSymbols = Nothing
    Set e3App = Nothing
End Sub

' === Основной запуск ===
Dim e3App
Set e3App = CreateObject("CT.Application")

e3App.PutInfo 0, "=== СТАРТ СКРИПТА ОЧИСТКИ АТРИБУТОВ POO СИМВОЛОВ ==="
Call ClearAllPOOSymbolsAttributes()
e3App.PutInfo 0, "=== КОНЕЦ СКРИПТА ==="

Set e3App = Nothing