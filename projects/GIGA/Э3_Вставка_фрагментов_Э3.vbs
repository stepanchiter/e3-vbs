'*******************************************************************************
' Название скрипта: E3_PlaceFragmentsSorted
' Автор: E3.series VBScript Assistant
' Дата: 15.07.2025
' Описание: Скрипт для автоматического размещения фрагментов на листах Э3.
'          Сначала находит все символы "OOO", сортирует их по индексу,
'          а затем вставляет фрагменты в отсортированном порядке.
'*******************************************************************************
Option Explicit

Sub PlaceFragmentsOnSheets()
    Dim result
    result = MsgBox("Начать размещение фрагментов на листах Э3?", vbOKCancel + vbQuestion, "Подтверждение")
    
    If result = vbCancel Then
        Exit Sub
    End If
    
    Dim e3App, job, symbol, sheet, device
    
    ' --- Объявления основных переменных ---
    Dim allSymbolIds
    Dim allSymbolCount
    Dim currentSymbolId, symbolName, symbolIndex
    Dim subcircuitPath, subcircuitVersion
    Dim attrValue
    
    ' --- Переменные для работы с листами ---
    Dim currentSheetId
    Dim fragmentCountOnCurrentSheet
    Dim sheetNumber
    Dim sheetName
    Dim allSheetIds
    Dim sheetCount
    Dim foundSheet
    Dim i
    
    ' --- Переменные для работы с девайсами (устройствами) ---
    Dim existingDeviceIdsBeforeInsert
    Dim existingDeviceCountBeforeInsert
    Dim existingDevicesMap
    Dim allDeviceIdsAfterInsert
    Dim allDeviceCountAfterInsert
    Dim newDeviceIds()
    Dim newDeviceCount
    Dim currentDeviceId
    Dim devName, newName
    Dim prefixList, p, prefix
    
    ' --- Координаты для вставки фрагментов ---
    Dim xCoords(5)
    Dim yCoords(5)    
    Dim insertX, insertY
    
    ' --- Переменная 'd' является счетчиком в нескольких циклах ---
    Dim d
    
    ' --- Переменные для сортировки ---
    Dim oooSymbols()
    Dim oooSymbolCount
    Dim currentOOOId
    Dim currentOOOIndex
    Dim tempSymbolId
    Dim tempSymbolIndex
    Dim j
    
    xCoords(1) = 36
    yCoords(1) = 24
    
    xCoords(2) = 108
    yCoords(2) = 24
    
    xCoords(3) = 180
    yCoords(3) = 24
    
    xCoords(4) = 252
    yCoords(4) = 24
    
    xCoords(5) = 324
    yCoords(5) = 24

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    Set sheet = job.CreateSheetObject()
    Set device = job.CreateDeviceObject()    

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Размещение фрагментов ==="

    ' Инициализация переменных для листов
    currentSheetId = 0
    fragmentCountOnCurrentSheet = 0
    sheetNumber = 1
    oooSymbolCount = 0

    ' --- НАЧАЛО МОДИФИКАЦИИ: Сбор и сортировка "OOO" символов ---
    e3App.PutInfo 0, "Поиск и сортировка символов 'OOO'..."

    allSymbolCount = job.GetSymbolIds(allSymbolIds)
    ReDim oooSymbols(1, -1)

    If allSymbolCount > 0 Then
        For i = 1 To allSymbolCount
            currentSymbolId = allSymbolIds(i)
            symbol.SetId(currentSymbolId)
            symbolName = symbol.GetName()

            If UCase(Left(symbolName, 3)) = "OOO" Then
                symbolIndex = CLng(Mid(symbolName, 4))
                
                oooSymbolCount = oooSymbolCount + 1
                ReDim Preserve oooSymbols(1, oooSymbolCount - 1)
                oooSymbols(0, oooSymbolCount - 1) = currentSymbolId
                oooSymbols(1, oooSymbolCount - 1) = symbolIndex
            End If
        Next
    End If

    If oooSymbolCount = 0 Then
        e3App.PutInfo 0, "В проекте не найдено символов 'OOO'."
    Else
        ' Сортировка массива oooSymbols по индексу
        For i = 0 To oooSymbolCount - 2
            For j = i + 1 To oooSymbolCount - 1
                If oooSymbols(1, i) > oooSymbols(1, j) Then
                    ' Меняем местами ID
                    tempSymbolId = oooSymbols(0, i)
                    oooSymbols(0, i) = oooSymbols(0, j)
                    oooSymbols(0, j) = tempSymbolId
                    
                    ' Меняем местами индексы
                    tempSymbolIndex = oooSymbols(1, i)
                    oooSymbols(1, i) = oooSymbols(1, j)
                    oooSymbols(1, j) = tempSymbolIndex
                End If
            Next
        Next
        e3App.PutInfo 0, "Символы 'OOO' успешно отсортированы."

        ' --- НАЧАЛО ОСНОВНОГО ЦИКЛА ПО ОТСОРТИРОВАННЫМ СИМВОЛАМ ---
        For i = 0 To oooSymbolCount - 1
            currentSymbolId = oooSymbols(0, i)
            symbolIndex = oooSymbols(1, i)
            
            symbol.SetId(currentSymbolId)
            symbolName = symbol.GetName()
            
            attrValue = Trim(symbol.GetAttributeValue("ОД D_Proizv3"))

            If attrValue <> "" And (attrValue >= "1" And attrValue <= "6") Then
                Select Case attrValue
                    Case "1": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\1.e3p"
                    Case "2": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\2.e3p"
					Case "3": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\3.e3p"
					Case "4": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\4.e3p"
					Case "5": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\5.e3p"
					Case "6": subcircuitPath = "C:\Users\SEK\Desktop\DWG_4_E3\Новая папка\ОДНОЛИН\фрагменты\6.e3p"
                End Select
                
                subcircuitVersion = "1"

                If fragmentCountOnCurrentSheet = 5 Then
                    fragmentCountOnCurrentSheet = 0
                    sheetNumber = sheetNumber + 1
                    currentSheetId = 0
                End If

                If currentSheetId = 0 Then
                    sheetName = CStr(sheetNumber)
                    
                    sheetCount = job.GetSheetIds(allSheetIds)
                    foundSheet = False
                    
                    If sheetCount > 0 Then
                        For d = 1 To sheetCount
                            sheet.SetId allSheetIds(d)
                            If sheet.GetName() = sheetName And UCase(Trim(sheet.GetAttributeValue("Код документа"))) = "Э3" Then
                                currentSheetId = allSheetIds(d)
                                foundSheet = True
                                e3App.PutInfo 0, "Найден лист для вставки: '" & sheetName & "' (ID: " & currentSheetId & ")"
                                Exit For
                            End If
                        Next
                    End If
                    
                    If Not foundSheet Then
                        e3App.PutInfo 2, "Ошибка: Лист '" & sheetName & "' с атрибутом 'Код документа' = 'Э3' не найден. Скрипт остановлен."
                        Exit For ' Выходим из цикла обработки символов
                    End If
                End If
                
                If foundSheet Or currentSheetId <> 0 Then
                    sheet.SetId currentSheetId
                    
                    fragmentCountOnCurrentSheet = fragmentCountOnCurrentSheet + 1
                    insertX = xCoords(fragmentCountOnCurrentSheet)
                    insertY = yCoords(fragmentCountOnCurrentSheet)
                    
                    e3App.PutInfo 0, "Обработка " & symbolName & " (атрибут = " & attrValue & ") > вставка: " & subcircuitPath & " на лист '" & sheet.GetName() & "' в позиции (" & insertX & ", " & insertY & ")"

                    existingDeviceCountBeforeInsert = job.GetDeviceIds(existingDeviceIdsBeforeInsert)            
                    Set existingDevicesMap = CreateObject("Scripting.Dictionary")            
                    If existingDeviceCountBeforeInsert > 0 Then
                        For d = 1 To existingDeviceCountBeforeInsert            
                            existingDevicesMap.Add existingDeviceIdsBeforeInsert(d), True            
                        Next
                    End If
                    e3App.PutInfo 0, "Количество устройств в проекте до вставки: " & existingDeviceCountBeforeInsert

                    Dim insertResult
                    insertResult = sheet.PlacePart(subcircuitPath, subcircuitVersion, insertX, insertY, 0.0)

                    If insertResult = 0 Or insertResult = -3 Then
                        e3App.PutInfo 0, "Фрагмент успешно установлен. Запуск переименования новых девайсов..."

                        allDeviceCountAfterInsert = job.GetDeviceIds(allDeviceIdsAfterInsert)            
                        e3App.PutInfo 0, "Количество устройств в проекте после вставки: " & allDeviceCountAfterInsert

                        newDeviceCount = 0
                        ReDim newDeviceIds(-1)

                        If allDeviceCountAfterInsert > 0 Then
                            For d = 1 To allDeviceCountAfterInsert            
                                currentDeviceId = allDeviceIdsAfterInsert(d)
                                If Not existingDevicesMap.Exists(currentDeviceId) Then
                                    If newDeviceCount = 0 Then
                                        ReDim newDeviceIds(0)
                                    Else
                                        ReDim Preserve newDeviceIds(newDeviceCount)
                                    End If
                                    newDeviceIds(newDeviceCount) = currentDeviceId
                                    newDeviceCount = newDeviceCount + 1
                                End If
                            Next
                        End If
                        e3App.PutInfo 0, "Найдено новых девайсов для переименования: " & newDeviceCount

                        prefixList = Array("-tQF", "-tKM")
                        
                        If newDeviceCount > 0 Then
                            For d = 0 To newDeviceCount - 1
                                currentDeviceId = newDeviceIds(d)
                                device.SetId currentDeviceId
                                devName = device.GetName()

                                For p = 0 To UBound(prefixList)
                                    prefix = prefixList(p)
                                    If LCase(Left(devName, Len(prefix))) = LCase(prefix) Then
                                        newName = prefix & symbolIndex
                                        e3App.PutInfo 0, "Переименование девайса: '" & devName & "' > '" & newName & "'"
                                        device.SetName newName
                                        Exit For
                                    End If
                                Next
                            Next
                        Else
                            e3App.PutInfo 1, "Внимание: Не удалось найти новые девайсы для переименования после вставки фрагмента."
                        End If
                        Set existingDevicesMap = Nothing

                    Else
                        e3App.PutInfo 2, "Ошибка вставки фрагмента (код: " & insertResult & ") для символа " & symbolName
                    End If
                End If
            Else
                e3App.PutInfo 1, "Пропуск " & symbolName & ": некорректное значение атрибута 'ОД D_Proizv3' = '" & attrValue & "'"
            End If
        Next
    End If
    e3App.PutInfo 0, "=== СКРИПТ ЗАВЕРШЕН ==="

    Set device = Nothing
    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

Call PlaceFragmentsOnSheets()