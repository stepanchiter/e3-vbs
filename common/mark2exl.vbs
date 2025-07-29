' Глобальные переменные для массивов
Dim A(), A_Names(), B(), B_Names(), C(), C_Names(), E(), E_Names()

Set app = CreateObject( "CT.Application" ) 
Set job = app.CreateJobObject()
Set device = job.CreateDeviceObject()
Set conductor = job.CreatePinObject()
Set pin = job.CreatePinObject()
Set dev = job.CreateDeviceObject()
Set Sig = Job.CreateSignalObject()
set Cab = Job.CreateDeviceObject()
set Cor = Job.CreatePinObject()

' Формируем имя файла автоматически
Dim excelFilePath
excelFilePath = GetExcelFileName(Job)
If excelFilePath = "" Then
    WScript.Quit
End If

'call SheetGrid
'call IndexSet
'call Sort
'call ReNameWire
'call Sort

' Инициализация массивов перед использованием
ReDim A(0) : ReDim A_Names(0)
ReDim B(0) : ReDim B_Names(0)
ReDim C(0) : ReDim C_Names(0)
ReDim E(0) : ReDim E_Names(0)


call GetWireCrossSections
call ExportWireDataToExcel(excelFilePath, A, A_Names, B, B_Names, C, C_Names, E, E_Names)

Sub IndexSet ' имя провода = имя цепи
	deviceCount = Job.GetAllDeviceIds( deviceIds )     
	If deviceCount > 0 Then 
		For deviceIndex = 1 To deviceCount
		deviceId = device.SetId( deviceIds( deviceIndex ) )
		conductorCount = device.GetAllCoreIds( conductorIds )
			If conductorCount > 0 Then
			deviceName = device.GetName()
				if deviceName = "Провода" then
					For conductorIndex = 1 To conductorCount
						conductorId = conductor.SetId( conductorIds( conductorIndex ) )				
						result = conductor.GetEndPinId( 1 )
						pinId = pin.SetId( result )
						pinName = pin.GetName()
						devId = dev.SetId( pinId )
						devName = dev.GetName()
						signal = pin.GetSignalName()
						conductor.SetName(signal)
					
					next
				end if
			end if
		next
	end If 
end Sub



sub Sort ' сортируем провода в дереве изделий
	Const ASCENDING = 1    ' по возрастанию	
	' Const DESCENDING = 2 ' по убыванию
	Dim sortOrder : sortOrder = ASCENDING
 
	deviceCount = Job.GetAllDeviceIds( deviceIds )
	If deviceCount > 0 Then 
		For deviceIndex = 1 To deviceCount
			deviceId = device.SetId( deviceIds( deviceIndex ) )
			deviceName = device.GetName()
			if deviceName = "Провода" then
			result = device.Sort( sortOrder )
				If result = 0 Then
					message = "Error sorting device " & deviceName & " ( " & deviceId & " )"
				Else
					message = "Изделие: " & deviceName & " ( " & deviceId & " ) отсортированы."
				End If
				app.PutInfo 0, "==========================================================="
				app.PutInfo 0, message    
			end if
		Next
	End If
end sub



Sub SheetGrid ' имя цепей по позиции
	sigcnt = Job.GetSignalIds(sigids)
	If sigcnt = 0 Then
		App.PutInfo 1, "No signals found, exiting..."
		WScript.Quit
	End If

' Формат для переименования
	const FORMAT = "#<.SHEET><.GRID>"

' Переименовываем цепи, начинающиеся с #
	For i = 1 To sigcnt
		Sig.SetId sigids(i)
		SignalName = Sig.GetName
		If Left(SignalName, 1) = "#" Then
			result = Sig.SetNameFormat(FORMAT)
			If result = 0 Then
				App.PutInfo 1, "ошибка переименования цепи " & SignalName
			Else
				App.PutInfo 0, "цепь " & SignalName & " переименована в формат " & FORMAT
			End If
		End If
	Next
end sub


Sub ReNameWire ' имя провода по цвету
	nCabs = Job.GetCableCount				' Количество кабелей в проекте
	if nCabs = 0 then
		App.PutInfo 1, "No cables in project, exiting..."
		wscript.quit
	end If

	LCounter = 1
	L1Counter = 1
	L2Counter = 1
	L3Counter = 1
	L1aCounter = 1
	NCounter = 5

' Создаем словарь для хранения соответствия исходных и новых имен
	Set NameMapping = CreateObject("Scripting.Dictionary")

	cablecount = Job.GetCableIds(cableids)
	For i = 1 To cablecount
		Cab.SetId cableids(i)
		If Cab.IsWiregroup Then
			wircnt = Cab.GetPinIds(wirids)
			For j = 1 To wircnt
				Cor.SetId wirids(j)							
				WireName = Cor.GetName
				WireColor = Cor.GetColourDescription
				' 1. Переименовываем только провода, имя которых начинается с #
					If Left(WireName, 1) = "#" Then
					' Проверяем, есть ли уже новое имя для этого исходного имени
						If NameMapping.Exists(WireName) Then
							NewName = NameMapping(WireName)
						Else
						' 2. Черный провод
							If WireColor = "черный" Then
								NewName = "L1." & L1Counter
								L1Counter = L1Counter + 1
						' 3. Коричневый провод
							ElseIf WireColor = "коричневый" Then
								NewName = "L2." & L2Counter
								L2Counter = L2Counter + 1
						' 4. Серый провод
							ElseIf WireColor = "серый" Then
								NewName = "L3." & L3Counter
								L3Counter = L3Counter + 1
						' 5. Синий провод
							ElseIf WireColor = "синий" Then
								NewName = "N" & NCounter
								NCounter = NCounter + 1	
							
							Else
								NewName = LCounter
								LCounter = LCounter + 1	
							End If

						' Сохраняем новое имя в словаре
							NameMapping.Add WireName, NewName
						End If

					' Переименовываем провод
						If Not IsEmpty(NewName) Then
							Cor.SetName NewName
							App.PutInfo 0, "провод " & WireName & " переименован в " & NewName
						End If
					End If
			Next
		End If 
	Next

end sub




Sub GetWireCrossSections
    nCabs = Job.GetCableCount
    If nCabs = 0 Then
        App.PutInfo 1, "No cables in project, exiting..."
        wscript.quit
    End If

    ' Массивы для хранения данных
    Dim WireCrossSections(), WireNames()
    ReDim WireCrossSections(0)
    ReDim WireNames(0)

    ' Сбор данных
    cablecount = Job.GetCableIds(cableids)
    For i = 1 To cablecount
        Cab.SetId cableids(i)
        If Cab.IsWiregroup Then
            wircnt = Cab.GetPinIds(wirids)
            For j = 1 To wircnt
                Cor.SetId wirids(j)
                WireCrossSection = Cor.GetCrossSectionDescription
                WireName = Cor.GetName
                
                If WireCrossSection <> "" Then
                    If UBound(WireCrossSections) = 0 And WireCrossSections(0) = "" Then
                        WireCrossSections(0) = WireCrossSection
                        WireNames(0) = WireName
                    Else
                        ReDim Preserve WireCrossSections(UBound(WireCrossSections) + 1)
                        ReDim Preserve WireNames(UBound(WireNames) + 1)
                        WireCrossSections(UBound(WireCrossSections)) = WireCrossSection
                        WireNames(UBound(WireNames)) = WireName
                    End If
                End If
            Next
        End If
    Next

    For k = 0 To UBound(WireCrossSections)
        section = WireCrossSections(k)
        wireName = WireNames(k)
        
        If InStr(section, "0.75") > 0 Or InStr(section, "1.5") > 0 Then
            AddToArray A, A_Names, section, wireName
        ElseIf InStr(section, "10") > 0 Or InStr(section, "16") > 0 Or InStr(section, "25") > 0 Or InStr(section, "35") > 0 Then
            AddToArray C, C_Names, section, wireName
        ElseIf InStr(section, "2.5") > 0 Or InStr(section, "4") > 0 Or InStr(section, "6") > 0 Then
            AddToArray B, B_Names, section, wireName
        ElseIf InStr(section, "50") > 0 Or InStr(section, "70") > 0 Then
            AddToArray E, E_Names, section, wireName
        End If
    Next

    ' Вывод информации
    ShowGroupInfo "A", A, A_Names, "0,75-1,5", "ТМАРК-НГ-2П-3,2/1,6Б (50М), арт. КНГ2П-032-Б50"
    ShowGroupInfo "B", B, B_Names, "2,5-6", "ТМАРК-НГ-2П-6,4/3,2Б (50М), арт. КНГ2П-064-Б50"
    ShowGroupInfo "C", C, C_Names, "10-35", "ТМАРК-НГ-2П-12,7/6,4Б (50М), арт. КНГ2П-127-Б50"
    ShowGroupInfo "E", E, E_Names, "50-70", "ТМАРК-НГ-2П-19,1/9,5Б (50М), арт. КНГ2П-191-Б50"
End Sub

' Отдельная процедура для добавления в массив
Sub AddToArray(arr(), namesArr(), section, wireName)
    If UBound(arr) = 0 And arr(0) = "" Then
        arr(0) = section
        namesArr(0) = wireName
    Else
        ReDim Preserve arr(UBound(arr) + 1)
        ReDim Preserve namesArr(UBound(namesArr) + 1)
        arr(UBound(arr)) = section
        namesArr(UBound(namesArr)) = wireName
    End If
End Sub

' Отдельная процедура для вывода информации
Sub ShowGroupInfo(groupName, sectionsArray, namesArray, sizeRange, tubeSpec)
    If UBound(sectionsArray) >= 0 And sectionsArray(0) <> "" Then
        tubeLength = (UBound(sectionsArray) + 1) * 2 * 0.015
        App.PutInfo 0, "-------------------------------"
        App.PutInfo 0, "Группа " & groupName & " (сечение " & sizeRange & "):"
        App.PutInfo 0, "Трубка: " & tubeSpec
        App.PutInfo 0, "Общая длина: " & tubeLength & " м"
        App.PutInfo 0, "Список проводов:"
        
        For m = 0 To UBound(namesArray)
            App.PutInfo 0, "  " & namesArray(m) & " (" & sectionsArray(m) & ")"
        Next
    End If
End Sub

' Вспомогательная процедура для добавления данных в массивы
Sub AddToArray(arr(), namesArr(), section, wireName)
    If UBound(arr) = 0 And arr(0) = "" Then
        arr(0) = section
        namesArr(0) = wireName
    Else
        ReDim Preserve arr(UBound(arr) + 1)
        ReDim Preserve namesArr(UBound(namesArr) + 1)
        arr(UBound(arr)) = section
        namesArr(UBound(namesArr)) = wireName
    End If
End Sub


'==============================================
' Процедура формирования имени XLS файла
' Возвращает полный путь к файлу в формате:
' <путь_проекта>\<имя_проекта>.xls
' с заменой "Sch2_" на "brk" и удалением ".e3d"
'==============================================
Function GetExcelFileName(Job)
    Dim projectPath, projectName, excelFileName
    
    ' Получаем путь проекта
    projectPath = Job.GetPath()
    If Len("" & projectPath) = 0 Then
        App.PutInfo 1, "Ошибка получения пути проекта"
        GetExcelFileName = ""
        Exit Function
    End If
    
    ' Получаем имя проекта
    projectName = Job.GetName()
    If Len("" & projectName) = 0 Then
        App.PutInfo 1, "Ошибка получения имени проекта"
        GetExcelFileName = ""
        Exit Function
    End If
    
    ' Заменяем "Sch2_" на "brk" в имени проекта
    projectName = Replace(projectName, "Sch2_", "brk", 1, -1, vbTextCompare)
    
    ' Убираем расширение .e3d если есть (регистронезависимо)
    projectName = Replace(projectName, ".e3d", "", 1, -1, vbTextCompare)
    projectName = Replace(projectName, ".E3D", "", 1, -1, vbTextCompare)
    
    ' Создаем имя файла Excel (без лишних суффиксов)
    excelFileName = projectPath & "\" & projectName & ".xls"
    
    App.PutInfo 0, "Сформирован путь к файлу Excel: " & excelFileName
    GetExcelFileName = excelFileName
End Function

'процедура для экспорта данных в Excel
Sub ExportWireDataToExcel(excelFilePath, A, A_Names, B, B_Names, C, C_Names, E, E_Names)
    On Error Resume Next
    
    ' Создаем объект Excel
    Dim ExcelApp, ExcelBook, ExcelSheet
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = True ' Для отладки
    
    ' Пытаемся открыть файл
    Set ExcelBook = ExcelApp.Workbooks.Open(excelFilePath)
    If Err.Number <> 0 Then
        App.PutInfo 1, "Ошибка открытия файла " & excelFilePath & ": " & Err.Description
        ExcelApp.Quit
        Set ExcelApp = Nothing
        Exit Sub
    End If
    
    ' Пытаемся получить лист
    Set ExcelSheet = Nothing
    On Error Resume Next
    Set ExcelSheet = ExcelBook.Sheets("МаркировкиПроводов")
    If ExcelSheet Is Nothing Then
        ' Создаем новый лист, если не существует
        Set ExcelSheet = ExcelBook.Sheets.Add
        ExcelSheet.Name = "МаркировкиПроводов"
    End If
    On Error GoTo 0
    
    ' Очищаем лист полностью
    ExcelSheet.Cells.Clear
    
    ' Устанавливаем ТЕКСТОВЫЙ формат для всех ячеек (новое!)
    ExcelSheet.Cells.NumberFormat = "@"
    
    ' Записываем шапку таблицы
    With ExcelSheet
        ' Диапазоны сечений
        .Cells(1, 1).Value = "0.75-1.5 мм2"
        .Cells(1, 2).Value = "2.5-6 мм2"
        .Cells(1, 3).Value = "10-35 мм2"
        .Cells(1, 4).Value = "50-70 мм2"
        
        ' Форматирование шапки (жирный, центрирование, цвет)
        With .Range("A1:D1")
            .Font.Bold = True
            .HorizontalAlignment = -4108 ' xlCenter
            .Interior.Color = 13434879 ' Светло-голубой фон
            .Borders.Weight = 2 ' xlThin
        End With
        
        ' Записываем данные по группам (все значения будут как текст)
        If UBound(A) >= 0 And A(0) <> "" Then
            For i = 0 To UBound(A)
                .Cells(i + 2, 1).Value = "'" & A_Names(i) ' Апостроф для гарантии текстового формата
            Next
        End If
        
        If UBound(B) >= 0 And B(0) <> "" Then
            For i = 0 To UBound(B)
                .Cells(i + 2, 2).Value = "'" & B_Names(i)
            Next
        End If
        
        If UBound(C) >= 0 And C(0) <> "" Then
            For i = 0 To UBound(C)
                .Cells(i + 2, 3).Value = "'" & C_Names(i)
            Next
        End If
        
        If UBound(E) >= 0 And E(0) <> "" Then
            For i = 0 To UBound(E)
                .Cells(i + 2, 4).Value = "'" & E_Names(i)
            Next
        End If
        
        ' Автоподбор ширины столбцов
        .Columns("A:D").AutoFit
    End With
    
    ' Сохраняем и закрываем
    ExcelBook.Save
    ExcelBook.Close
    ExcelApp.Quit
    
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    
    App.PutInfo 0, "Данные успешно экспортированы в " & excelFilePath & " на лист 'МаркировкиПроводов'"
End Sub

Set dev = Nothing
Set pin = Nothing
Set conductor = Nothing
Set device = Nothing   
Set job = Nothing 
Set app = Nothing
Set Sig = Nothing
Set Cab = Nothing
Set Cor = Nothing