'*******************************************************************************
' Название скрипта: E3_SynchronizeOooPooAttributes
' Автор: E3.series VBScript Assistant
' Дата: 08.07.2025
' Описание: Скрипт для синхронизации определенных атрибутов между символами "OOO" и "POO" в проекте E3.series.
'           Он ищет пары символов по числовому индексу в их имени и копирует значения атрибутов из "OOO" в "POO" символы,
'           если значения не пусты. Добавлено копирование атрибутов "ОД E_Iras" и "ОД E_Inom".
'*******************************************************************************
Option Explicit

Sub SynchronizeOooPooAttributes()
    ' Показываем предупреждающее окно
    Dim result
    result = MsgBox("Начать синхронизацию атрибутов между OOO и POO символами в проекте?", vbOKCancel + vbQuestion, "Подтверждение")
    
    ' Если пользователь нажал "Отмена", выходим из скрипта
    If result = vbCancel Then
        Exit Sub
    End If

    Dim e3App, job, oooSymbol, pooSymbol ' Раздельные объекты для OOO и POO символов
    Dim allSymbolIds ' Массив всех ID символов в проекте
    Dim allSymbolCount
    Dim s_ooo, s_poo ' Счетчики циклов
    Dim oooSymbolId, oooSymbolName, oooSymbolIndex
    Dim pooSymbolId, pooSymbolName
    Dim pooSymbolFound

    Dim attrNamesToCopy ' Массив с названиями атрибутов для копирования
    Dim currentAttrName ' Текущее название атрибута
    Dim oooAttrValue ' Значение атрибута из OOO символа
    Dim p ' Счетчик для цикла по атрибутам

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set oooSymbol = job.CreateSymbolObject()
    Set pooSymbol = job.CreateSymbolObject() ' Отдельный объект для POO символов

    e3App.PutInfo 0, "=== СТАРТ СКРИПТА: Синхронизация атрибутов OOO и POO ==="

    ' Обновленный список атрибутов для копирования, включая "ОД D_Proizv2", "ОД E_Iras", "ОД E_Inom"
    attrNamesToCopy = Array("ОД E_TAG", "ОД E_TYPE", "ОД E_Pnom", "ОД V_Type", "ОД D_Proizv2", "ОД E_Iras", "ОД E_Inom")

    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount > 0 Then
        For s_ooo = 1 To allSymbolCount ' Цикл по всем символам для поиска OOO
            oooSymbolId = allSymbolIds(s_ooo)
            oooSymbol.SetId(oooSymbolId)
            oooSymbolName = oooSymbol.GetName()

            ' Проверяем, начинается ли имя символа с "OOO"
            If UCase(Left(oooSymbolName, 3)) = "OOO" Then
                oooSymbolIndex = Mid(oooSymbolName, 4) ' Извлекаем индекс
                pooSymbolName = "POO" & oooSymbolIndex ' Формируем имя соответствующего POO символа
                
                pooSymbolId = 0 ' Сброс ID POO символа
                pooSymbolFound = False

                ' Ищем соответствующий POO символ среди ВСЕХ символов в проекте
                For s_poo = 1 To allSymbolCount
                    If allSymbolIds(s_poo) <> oooSymbolId Then ' Избегаем сравнения с самим OOO символом
                        pooSymbol.SetId(allSymbolIds(s_poo))
                        If UCase(pooSymbol.GetName()) = UCase(pooSymbolName) Then
                            pooSymbolId = allSymbolIds(s_poo)
                            pooSymbolFound = True
                            e3App.PutInfo 0, "Найден соответствующий POO символ: " & pooSymbolName & " (для OOO: " & oooSymbolName & ")"
                            Exit For ' POO символ найден, можно выйти из этого внутреннего цикла
                        End If
                    End If
                Next

                If pooSymbolFound Then
                    ' Копируем атрибуты
                    For p = 0 To UBound(attrNamesToCopy)
                        currentAttrName = attrNamesToCopy(p)
                        oooAttrValue = oooSymbol.GetAttributeValue(currentAttrName)
                        
                        If oooAttrValue <> "" Then ' Копируем только непустые значения
                            pooSymbol.SetAttributeValue currentAttrName, oooAttrValue
                            e3App.PutInfo 0, "  -> Скопирован атрибут '" & currentAttrName & "': '" & oooAttrValue & "'"
                        Else
                            e3App.PutInfo 1, "  -> Атрибут '" & currentAttrName & "' для символа " & oooSymbolName & " пустой. Пропущен."
                        End If
                    Next
                Else
                    e3App.PutInfo 1, "Внимание: Соответствующий POO символ '" & pooSymbolName & "' для OOO символа '" & oooSymbolName & "' не найден. Атрибуты не скопированы."
                End If
            End If ' End If для OOO символа
        Next
    Else
        e3App.PutInfo 0, "В проекте нет символов для обработки."
    End If

    e3App.PutInfo 0, "=== СКРИПТ ЗАВЕРШЕН ==="

    ' Очистка объектов
    Set oooSymbol = Nothing
    Set pooSymbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

Call SynchronizeOooPooAttributes()