Option Explicit

Sub ReplaceAllOOOWithSubcircuit()
    ' Показываем предупреждающее окно
    Dim result
    result = MsgBox("Заполняем всё фрагментами?", vbOKCancel + vbQuestion, "Подтверждение")
    
    ' Если пользователь нажал "Отмена", выходим из скрипта
    If result = vbCancel Then
        Exit Sub
    End If
    
    Dim e3App, job, symbol, sheet, device
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName, symbolIndex
    Dim s
    Dim subcircuitPath, subcircuitVersion
    Dim insertedDeviceIds(), deviceCount, d, devName, newName
    Dim prefixList, p, prefix
    Dim attrValue

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    Set sheet = job.CreateSheetObject()
    Set device = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА ==="

    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount > 0 Then
        For s = 1 To allSymbolCount
            currentSymbolId = allSymbolIds(s)
            symbol.SetId(currentSymbolId)
            symbolName = symbol.GetName()

            If UCase(Left(symbolName, 3)) = "OOO" Then
                symbolIndex = Mid(symbolName, 4)
                attrValue = Trim(symbol.GetAttributeValue("ОД D_Proizv3"))

                If attrValue <> "" And (attrValue >= "1" And attrValue <= "6") Then

                    ' Выбор пути к фрагменту по значению атрибута
                    Select Case attrValue
                        Case "1": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\1_tQF3P.e3p"
                        Case "2": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\2_tQF3PtKM.e3p"
                        Case "3": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\3_tАЗД3PtKM.e3p"
                        Case "4": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\4_tАЗД2PtKM.e3p"
                        Case "5": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\5_tQFd2PtKM.e3p"
                        Case "6": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\6_tQF1P.e3p"
                    End Select

                    subcircuitVersion = "1"

                    Dim oooX, oooY, oooSheetId, gridDesc, colVal, rowVal
                    oooSheetId = symbol.GetSchemaLocation(oooX, oooY, gridDesc, colVal, rowVal)

                    If oooSheetId > 0 Then
                        sheet.SetId oooSheetId
                        e3App.PutInfo 0, "Обработка " & symbolName & " (атрибут = " & attrValue & ") > вставка: " & subcircuitPath

                        Dim insertResult : insertResult = sheet.PlacePart(subcircuitPath, subcircuitVersion, oooX, oooY, 0.0)

                        If insertResult = 0 Or insertResult = -3 Then
                            e3App.PutInfo 0, "Фрагмент успешно установлен"

                            ' Переименование девайсов после вставки
                            deviceCount = job.GetDeviceIds(insertedDeviceIds)
                            For d = 1 To deviceCount
                                device.SetId insertedDeviceIds(d)
                                devName = device.GetName()

                                prefixList = Array("-tQF", "-tKM", "-tKL")
                                For p = 0 To UBound(prefixList)
                                    prefix = prefixList(p)
                                    If LCase(Left(devName, Len(prefix))) = LCase(prefix) Then
                                        newName = Replace(prefix, "-t", "-") & symbolIndex
                                        e3App.PutInfo 0, "Переименование: " & devName & " > " & newName
                                        device.SetName newName
                                        Exit For
                                    End If
                                Next
                            Next
                        Else
                            e3App.PutInfo 0, "Ошибка вставки фрагмента (код: " & insertResult & ")"
                        End If
                    Else
                        e3App.PutInfo 0, symbolName & " не размещён на схеме"
                    End If
                Else
                    e3App.PutInfo 0, "Пропуск " & symbolName & ": некорректное значение атрибута 'ОД D_Proizv3' = '" & attrValue & "'"
                End If
            End If
        Next
    Else
        e3App.PutInfo 0, "В проекте нет символов."
    End If

    ' Очистка
    Set device = Nothing
    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

Call ReplaceAllOOOWithSubcircuit()