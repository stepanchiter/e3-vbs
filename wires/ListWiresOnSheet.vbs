                    App.PutInfo 0, "Атрибут 'Класс соединения (Компоновка)' установлен для coreID: " & coreIds(j)
Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject()
Set Device = Job.CreateDeviceObject()
Set Sheet = Job.CreateSheetObject()
Set Pin = Job.CreatePinObject()

' Объявляем массивы
Dim sheetIds()
Dim netIds()
Dim coreIds()

' Создаём объект tree для работы с деревом проекта
Dim Tree
Set Tree = Job.CreateTreeObject
' Устанавливаем активное дерево текущим
TreeId = Tree.SetId(Job.GetActiveTreeId())

' Поиск выделенных в дереве листов
Dim sheetCount : sheetCount = Tree.GetSelectedSheetIds(sheetIds)

If sheetCount = 0 Then
    ' Если нет выделенных листов, берем активный
    sheetCount = 1
    ReDim sheetIds(1)
    sheetIds(1) = Job.GetActiveSheetId()
    
    If sheetIds(1) = 0 Then
        App.PutInfo 1, "Нет активного или выделенного листа"
        WScript.Quit
    End If
    
    App.PutInfo 0, "Работаем с активным листом..."
Else
    App.PutInfo 0, "Работаем с выделенными листами..."
End If

' Перебираем все выбранные листы
For selectedNum = 1 To sheetCount
    ' Устанавливаем текущий лист
    Sheet.SetId sheetIds(selectedNum)
    Dim selectedSheetName : selectedSheetName = Sheet.GetName()
    App.PutInfo 0, "Обработка листа: " & selectedSheetName

' Получаем цепи на листе
Dim netCount : netCount = Sheet.GetNetIds(netIds)
If netCount = 0 Then ReDim netIds(0)
App.PutInfo 0, "=============== Провода на листе " & selectedSheetName & " ==============="

If netCount = 0 Then
    App.PutInfo 1, "На листе нет цепей"
Else
    ' Создаем объект для работы с цепями
    Dim Net : Set Net = Job.CreateNetObject()
    Dim wireCount : wireCount = 0

    ' Перебираем все цепи
    For i = 1 To netCount
        Net.SetId netIds(i)
        
        ' Получаем все жилы в цепи
        Dim coreCount : coreCount = Net.GetCoreIds(coreIds)
        If coreCount = 0 Then ReDim coreIds(0)
        If coreCount > 0 Then
            ' Перебираем жилы
            For j = 1 To coreCount
                Pin.SetId coreIds(j)
                Device.SetId coreIds(j)
                
                ' Если это провод
                If Device.IsWiregroup() Then
                    wireCount = wireCount + 1
                    ' Получаем информацию о проводе
                    Dim wireName : wireName = Device.GetName()
                    Dim wireId : wireId = Device.GetId()
                    Dim signalName : signalName = Pin.GetSignalName()
                    Dim colorDesc : colorDesc = Pin.GetColourDescription()
                    
                    ' Выводим информацию с кликабельным ID
                    App.PutMessageEx 0, wireCount & ". " & " (Цепь: " & signalName & ", Цвет: " & colorDesc & ", CoreID: " & coreIds(j) & ")", coreIds(j), 0, 0, 0
                        ' Устанавливаем атрибут "Класс соединения (Компоновка)" для core
                        Pin.SetAttributeValue "Класс соединения (Компоновка)", "230/400V"
                End If
            Next
        End If
    Next

    ' Очистка
    Set Net = Nothing

    ' Выводим итог
    If wireCount = 0 Then
        App.PutInfo 0, "На листе не найдено проводов"
    Else
        App.PutInfo 0, "==============================================="
        App.PutInfo 0, "Всего найдено проводов на листе " & selectedSheetName & ": " & wireCount
    End If
End If
Next ' Конец цикла по листам

' Очищаем объекты
Set Pin = Nothing
Set Sheet = Nothing
Set Device = Nothing
Set Job = Nothing
Set App = Nothing
