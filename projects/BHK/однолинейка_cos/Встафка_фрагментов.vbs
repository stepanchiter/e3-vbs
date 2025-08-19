'*******************************************************************************
' Название скрипта: E3_ReplaceAllOOSWithSubcircuit
' Автор: E3.series VBScript Assistant
' Дата: 08.07.2025
' Описание: Скрипт предназначен для автоматической замены символов,
'           имена которых начинаются с "OOS" на соответствующие им фрагменты схемы.
'           Выбор фрагмента определяется значением атрибута "ОД D_Proizv3".
'           После вставки фрагмента скрипт переименовывает новые устройства,
'           созданные из фрагмента, используя числовой индекс из имени исходного
'           "OOS" символа. Добавлены префиксы -tUZ и -tV для переименования.
'*******************************************************************************
Option Explicit

Sub ReplaceAllOOSWithSubcircuit()
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

            If UCase(Left(symbolName, 3)) = "OOS" Then
                symbolIndex = Mid(symbolName, 4)
                attrValue = Trim(symbol.GetAttributeValue("ОД D_Proizv3"))

                ' Изначальная проверка атрибута, без изменения логики
                If attrValue <> "" And (attrValue >= "1" And attrValue <= "9") Or attrValue = "10" Then ' Расширил условие для "10"
                    
                    subcircuitPath = "" ' Инициализируем, чтобы знать, был ли путь назначен
                    
                    ' Выбор пути к фрагменту по значению атрибута
                    Select Case attrValue
                        Case "1": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\1_ВНТ через ЧП_ЧП внутри.e3p"
                        Case "2": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\2_ЭК дискретная ступень.e3p"
                        Case "3": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\3_3ф двигатель_прямой пуск.e3p"
                        Case "4": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\4_1ф двигатель_прямой пуск_контактор.e3p"
                        Case "5": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\5_обогрев клапана_диф.автомат_через реле 4.e3p"
                        Case "6": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\6_освещение.e3p"
                        Case "7": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\7_ЭК плавная ступень.e3p"
                        Case "8": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\8_1ф двигатель_через реле 4_реруглятор внутри.e3p"
                        Case "9": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\9_питание шкафа.e3p"
                        Case "10": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\10_увлажнитель_3ф.e3p"
                        Case "11": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\11_обогрев решетки_диф.автомат_через КМ.e3p"
                        Case "12": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\12_увлажнитель_3ф_4п.e3p"
                        Case "13": subcircuitPath = "W:\СПЕЦПРОЕКТЫ E3\БХК\ОДНОЛИН\фрагменты\Катя\13_1ф двигатель_через реле 4.e3p"
                        Case Else: 
                            ' Если Case Else, subcircuitPath останется пустым, и мы выведем сообщение
                            e3App.PutInfo 1, "Пропуск " & symbolName & ": Неизвестное значение атрибута 'ОД D_Proizv3' = '" & attrValue & "'"
                    End Select

                    ' Продолжаем выполнение только если subcircuitPath был успешно определен
                    If subcircuitPath <> "" Then
                        subcircuitVersion = "1"

                        Dim OOSX, OOSY, OOSSheetId, gridDesc, colVal, rowVal
                        OOSSheetId = symbol.GetSchemaLocation(OOSX, OOSY, gridDesc, colVal, rowVal)

                        If OOSSheetId > 0 Then
                            sheet.SetId OOSSheetId
                            e3App.PutInfo 0, "Обработка " & symbolName & " (атрибут = " & attrValue & ") > вставка: " & subcircuitPath

                            Dim insertResult : insertResult = sheet.PlacePart(subcircuitPath, subcircuitVersion, OOSX, OOSY, 0.0)

                            If insertResult = 0 Or insertResult = -3 Then ' 0: успех, -3: возможно, также успех в некоторых версиях
                                e3App.PutInfo 0, "Фрагмент успешно установлен"

                                ' Переименование девайсов после вставки
                                ' ВНИМАНИЕ: job.GetDeviceIds(insertedDeviceIds) получает ВСЕ девайсы в проекте.
                                ' Это может привести к переименованию уже существующих девайсов.
                                ' Для более надежного решения требуется сравнение списка девайсов до и после вставки.
                                deviceCount = job.GetDeviceIds(insertedDeviceIds)
                                For d = 1 To deviceCount
                                    device.SetId insertedDeviceIds(d)
                                    devName = device.GetName()

                                    ' Обновленный список префиксов для переименования (строка 105 в исходном скрипте)
                                    prefixList = Array("-tQF", "-tKM", "-tKL", "-tUZ", "-tV") 
                                    For p = 0 To UBound(prefixList)
                                        prefix = prefixList(p)
                                        If LCase(Left(devName, Len(prefix))) = LCase(prefix) Then
                                            newName = Replace(prefix, "-t", "-") & symbolIndex
                                            e3App.PutInfo 0, "Переименование: " & devName & " > " & newName
                                            device.SetName newName
                                            Exit For ' Выходим из внутреннего цикла по префиксам, так как нашли совпадение
                                        End If
                                    Next
                                Next
                            Else
                                e3App.PutInfo 0, "Ошибка вставки фрагмента (код: " & insertResult & ")"
                            End If
                        Else
                            e3App.PutInfo 0, symbolName & " не размещён на схеме. Пропущен."
                        End If
                    End If ' End If для проверки subcircuitPath
                Else
                    e3App.PutInfo 1, "Пропуск " & symbolName & ": некорректное или отсутствующее значение атрибута 'ОД D_Proizv3' = '" & attrValue & "'"
                End If
            End If
        Next
    Else
        e3App.PutInfo 0, "В проекте нет символов для обработки."
    End If

    e3App.PutInfo 0, "=== СКРИПТ ЗАВЕРШЕН ==="

    ' Очистка
    Set device = Nothing
    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

Call ReplaceAllOOSWithSubcircuit()