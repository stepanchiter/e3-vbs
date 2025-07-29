Option Explicit

' =========================================================================
' Вспомогательная функция: Извлечение числового индекса из имени
' =========================================================================
Function ExtractNumericIndex(nameString)
    On Error Resume Next
    Dim i, char, numericPart
    numericPart = ""
    ExtractNumericIndex = 0

    If Len(nameString) = 0 Then Exit Function

    For i = Len(nameString) To 1 Step -1
        char = Mid(nameString, i, 1)
        If IsNumeric(char) Then
            numericPart = char & numericPart
        Else
            If Len(numericPart) > 0 Then Exit For
        End If
    Next

    If Len(numericPart) > 0 Then
        If IsNumeric(numericPart) Then
            ExtractNumericIndex = CInt(numericPart)
        End If
    End If

    If Err.Number <> 0 Then
        Err.Clear
    End If
End Function

' =========================================================================
' Основная процедура
' =========================================================================
Sub RenameTerminalBlocksBySystem()
    On Error Resume Next

    Dim e3App, job, device, symbol
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set device = job.CreateDeviceObject()
    Set symbol = job.CreateSymbolObject()

    Dim systemTerminalNameMap
    Set systemTerminalNameMap = CreateObject("Scripting.Dictionary")

    Dim pooIndexToSystemNameMap
    Set pooIndexToSystemNameMap = CreateObject("Scripting.Dictionary")

    Dim terminalNameCounter
    terminalNameCounter = 1

    Dim i, currentSymbolId, currentSymbolName
    Dim attrValue, systemName, dashPos, pooNumericIndex
    Dim allSymbolIds, symbolCount

    e3App.PutInfo 0, "=== СТАРТ: Поиск и переименование клеммников по системам ==="

    ' ФАЗА 1
    e3App.PutMessageEx 0, "Фаза 1: Сбор системных имен и генерация шаблонов клеммников -tXTN...", 0, 0, 0, 249

    symbolCount = job.GetSymbolIds(allSymbolIds)

    If symbolCount > 0 Then
        For i = 1 To symbolCount
            currentSymbolId = allSymbolIds(i)
            symbol.SetId(currentSymbolId)
            currentSymbolName = symbol.GetName()

            If Left(UCase(currentSymbolName), 3) = "POO" Then
                attrValue = symbol.GetAttributeValue("ОД E_TAG")
                If attrValue <> "" Then
                    dashPos = InStr(1, attrValue, "-")
                    If dashPos > 1 Then
                        systemName = Left(attrValue, dashPos - 1)
                    Else
                        systemName = "(E_TAG не имеет дефиса: '" & attrValue & "')"
                    End If
                Else
                    systemName = "(E_TAG пустой)"
                End If

                If Not systemTerminalNameMap.Exists(systemName) Then
                    systemTerminalNameMap.Add systemName, "-tXT" & terminalNameCounter
                    e3App.PutInfo 0, "Новая система: " & systemName & " -> " & systemTerminalNameMap(systemName)
                    terminalNameCounter = terminalNameCounter + 1
                End If
            End If
        Next
    End If

    ' ФАЗА 2
    e3App.PutMessageEx 0, "Фаза 2: Сбор данных POOi -> Система...", 0, 0, 0, 249

    If symbolCount > 0 Then
        For i = 1 To symbolCount
            currentSymbolId = allSymbolIds(i)
            symbol.SetId(currentSymbolId)
            currentSymbolName = symbol.GetName()

            If Left(UCase(currentSymbolName), 3) = "POO" Then
                pooNumericIndex = ExtractNumericIndex(currentSymbolName)
                If pooNumericIndex > 0 Then
                    attrValue = symbol.GetAttributeValue("ОД E_TAG")
                    If attrValue <> "" Then
                        dashPos = InStr(1, attrValue, "-")
                        If dashPos > 1 Then
                            systemName = Left(attrValue, dashPos - 1)
                        Else
                            systemName = "(E_TAG не имеет дефиса: '" & attrValue & "')"
                        End If
                    Else
                        systemName = "(E_TAG пустой)"
                    End If

                    If Not pooIndexToSystemNameMap.Exists(pooNumericIndex) Then
                        pooIndexToSystemNameMap.Add pooNumericIndex, systemName
                        e3App.PutInfo 0, "POO" & pooNumericIndex & " -> " & systemName
                    End If
                End If
            End If
        Next
    End If

    ' ФАЗА 3
    e3App.PutMessageEx 0, "Фаза 3: Переименование клеммников -XTN в -tXTN...", 0, 0, 0, 249

    Dim allTerminalDeviceIds, terminalDeviceCount
    Dim currentTerminalDeviceId, oldTerminalBlockName, xtNumericIndex, newTerminalBlockName, setResult

    terminalDeviceCount = job.GetTerminalIds(allTerminalDeviceIds)

    If terminalDeviceCount > 0 Then
        For i = 1 To terminalDeviceCount
            currentTerminalDeviceId = allTerminalDeviceIds(i)
            device.SetId(currentTerminalDeviceId)
            oldTerminalBlockName = device.GetName()

            If Left(UCase(oldTerminalBlockName), 3) = "-XT" Then
                xtNumericIndex = ExtractNumericIndex(oldTerminalBlockName)
                If xtNumericIndex > 0 Then
                    If pooIndexToSystemNameMap.Exists(xtNumericIndex) Then
                        systemName = pooIndexToSystemNameMap(xtNumericIndex)
                        If systemTerminalNameMap.Exists(systemName) Then
                            newTerminalBlockName = systemTerminalNameMap(systemName)
                            If StrComp(oldTerminalBlockName, newTerminalBlockName, vbTextCompare) <> 0 Then
                                setResult = device.SetName(newTerminalBlockName)
                                If setResult = 0 Then
                                    e3App.PutInfo 0, "Переименован: '" & oldTerminalBlockName & "' (ID: " & currentTerminalDeviceId & ") -> '" & newTerminalBlockName & "'"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If

    ' ФАЗА 4
    e3App.PutMessageEx 0, "Фаза 4: Переименование клеммников -tXTN в -XTN...", 0, 0, 0, 249

    terminalDeviceCount = job.GetTerminalIds(allTerminalDeviceIds)

    If terminalDeviceCount > 0 Then
        For i = 1 To terminalDeviceCount
            currentTerminalDeviceId = allTerminalDeviceIds(i)
            device.SetId(currentTerminalDeviceId)
            oldTerminalBlockName = device.GetName()

            If Left(UCase(oldTerminalBlockName), 4) = "-TXT" Then
                xtNumericIndex = ExtractNumericIndex(oldTerminalBlockName)
                If xtNumericIndex > 0 Then
                    newTerminalBlockName = "-XT" & xtNumericIndex
                    If StrComp(oldTerminalBlockName, newTerminalBlockName, vbTextCompare) <> 0 Then
                        setResult = device.SetName(newTerminalBlockName)
                        If setResult = 0 Then
                            e3App.PutInfo 0, "Финальное переименование: '" & oldTerminalBlockName & "' (ID: " & currentTerminalDeviceId & ") -> '" & newTerminalBlockName & "'"
                        End If
                    End If
                End If
            End If
        Next
    End If

    e3App.PutInfo 0, "=== ЗАВЕРШЕНО ==="

    Set systemTerminalNameMap = Nothing
    Set pooIndexToSystemNameMap = Nothing
    Set symbol = Nothing
    Set device = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

Call RenameTerminalBlocksBySystem()
