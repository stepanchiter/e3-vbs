' ===============================================================
' ===============================================================
'ред. 03.10.2024
' Скрипт предназначен для заполнения данных в Е3 из отчета F20
'скрипт берет данные из файла EXCEL
'путь указывается через меню
'скрипт отредактирован под пустые ячейки , есть возможность записывать из нескольких файлов F20

' ===============================================================
' ===============================================================



' Функция подключения к E3
'Set e3Application = CreateObject( "CT.Application" )
Set app = E3Connection()
Set job = app.CreateJobObject()
Set device = job.CreateDeviceObject()
Set dev = job.CreateDeviceObject()
Set symbol = job.CreateSymbolObject()
Set sym = job.CreateSymbolObject()
Set sheet = job.CreateSheetObject()
Set graphic = job.CreateGraphObject()
Set tree = job.CreateTreeObject()

Set Cab     = Job.CreateDeviceObject
Set Cor     = Job.CreatePinObject
Set Pin     = Job.CreatePinObject
Set Sheet   = Job.CreateSheetObject
Set Pin1     = Job.CreatePinObject
Set Pin2    = Job.CreatePinObject
Set Dev1     = Job.CreateDeviceObject
Set Dev2     = Job.CreateDeviceObject
Set signal = job.CreateSignalObject()
Set devicePin = job.CreatePinObject()


Set Excel = CreateObject("Excel.Application") 	' создаем объект Excel
Set objExcel = CreateObject("Excel.Application")

' Создание глобальных переменных
Dim app, appId
' ===============================================================
' Функция подключения к E3
' ===============================================================
Function E3Connection()
	' Улавливатель ошибок
	On Error Resume Next
		' Подключение процессу E3
		Set app = CreateObject("CT.Application")
		' Запрос идентификатора процесса
		appId = app.GetId()
		' Вывод сообщения об ошибке
		If (appId = 0) Then
			' Вывод сообщения
			MsgBox "Ошибка. Процесс E3.series не запущен или COM-классы приложения E3.series не зарегистрированы!", 16, "Ошибка"
		End If
	On Error Goto 0
	' Возврат функции
	Set E3Connection = app
End Function

' ===============================================================
' Функция работы с проектом
' ===============================================================
'Function E3Job(ByRef jobId)
	' Создание переменных
'	Dim job
'	Set job = app.CreateJobObject()
	' Запрос идентификатора проекта
'	jobId = job.GetId()
	' Проверка идентификатора
'	If (jobId = 0) Then
		' Вывод сообщения об ошибке
'		app.PutError 0, "Проект не открыт!"
		' Процедура завершения работы
'		Call ExitScript (False, job)
'	End If
	' Возврат функции
'	Set E3Job = job
'End Function


' ===============================================================
' Процедура выхода из работы скрипта
' ===============================================================
Sub ExitScript(ByVal flagSuccessExit, ByRef job) 
	' Проверка флага
	If (flagSuccessExit) Then
		' Вывод сообщения об успешном окончании
		Call app.PutInfo(0, "=====================================")
		Call app.PutInfo(0, "Автоматизация выполнена успешно!")
		Call app.PutInfo(0, "=====================================")
		
	Else
		' Вывод сообщения о не успешном окончании
		Call app.PutError(1, "Автоматизация не выполнена!")
	End If
	
	' Очистка объектов
	Set job = Nothing
	Set app = Nothing
	' Выход из скрипта
	WScript.Quit
End Sub




' ===============================================================
' ===============================================================
' Открытие файла "F20_R2_реверс в Е3.xlsx"
' ПОТОМ НАДО ПОДУМАТЬ, КАК УКАЗЫВАТЬ ЛЮБОЙ ФАЙЛ ОТЧЕТА
' ===============================================================
' ===============================================================

'fileName = "BGCC_16190-11-переченьIO_R4.xlsx"
'Dim fso
'Set fso = CreateObject("Scripting.FileSystemObject")
'	Dim thisFolder: thisFolder = fso.GetParentFolderName(WScript.ScriptFullName)
	' Полный путь до нужного файлами
'	Dim fileFullName
'	fileFullName = fso.BuildPath(thisFolder, fileName)
	' Проверка файла
'	If (fso.FileExists(fileFullName)) Then
'		objExcel.Visible = False 
'		objExcel.Visible = True 
'		objExcel.Workbooks.Open fileFullName
'	Else
		' Вывод сообщения об ошибке
'		Call MsgBox("Ошибка открытия файла " & fileFullName & ". Файла не существует!", 16, "Ошибка открытия файла")
		' Очистка объекта
'		Set fso = Nothing
		' Выход из выполнения скрипта
'		WScript.Quit
'	End If
	' Очистка объекта
'	Set fso = Nothing

'objExcel.Visible = False 
'objExcel.Visible = True 
'objExcel.Workbooks.Open "D:\E3_Generation\___.xlsx"
'objExcel.Worksheets("Лист1").Activate

'ii= 5



'--------------------------------------------------------------------------------------------

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
	Dim thisFolder: thisFolder = fso.GetParentFolderName(WScript.ScriptFullName)
	' Полный путь до нужного файлами
	Dim fileFullName
	fileFullName = InputBox("Введите путь к файлу EXCEL с модулями контроллера", "", "")
'	fileFullName = fso.BuildPath(thisFolder, fileName)
	' Проверка файла
	If (fso.FileExists(fileFullName)) Then
'		objExcel.Visible = False 
'		objExcel.Visible = True 
		objExcel.Workbooks.Open fileFullName
	Else
		' Вывод сообщения об ошибке
		Call MsgBox("Ошибка открытия файла " & fileFullName & ". Файла не существует!", 16, "Ошибка открытия файла")
		' Очистка объекта
		Set fso = Nothing
		' Выход из выполнения скрипта
		WScript.Quit
	End If
	' Очистка объекта
'	Set fso = Nothing

'objExcel.Visible = False 
objExcel.Visible = True 
'objExcel.Workbooks.Open "D:\E3_Generation\___.xlsx"
objExcel.Worksheets("Лист1").Activate

'--------------------------------------------------------------------------------------------















'			objExcel.Cells( i + 1 + ii+j, 1 ) = A1 													' НЕ ЗАПОЛНЯЕТСЯ (0) Номер строки / Raw No
'			objExcel.Cells( i + 1 + ii, 2 ) = ArrDeviceIds_2(i, 0) 													' (1) Позиционное обозначение / Position TAG
'			objExcel.Cells( i + 1 + ii, 4 ) = A1													' НЕ ЗАПОЛНЯЕТСЯ (2) Наименование сигнала / TAG
'			objExcel.Cells( i + 1 + ii, 6 ) = ArrDeviceIds_2(i, 1)													' (3) Описание сигнала / TAG discription
'			objExcel.Cells( i + 1 + ii, 9 ) = ArrDeviceIds_2(i, 5)													' (4) Внешнее подключение / Field connection 
'			objExcel.Cells( i + 1 + ii, 11 ) = ArrDeviceIds_2(i, 2)													' (5) Тип сигнала / Signal type
'			objExcel.Cells( i + 1 + ii, 12 ) = ArrDeviceIds_2(i, 3)													' (6) Тип подключения / Connection type
'			objExcel.Cells( i + 1 + ii, 13 ) = ArrDeviceIds_2(i, 4)													' (7) Ед. Изм / Measurement type
'			objExcel.Cells( i + 1 + ii+j, 14 ) = Dev.GetComponentAttributeValue("Тех. описание 1")	' НЕ ЗАПОЛНЯЕТСЯ(8) Тип модуля / Module type
'			objExcel.Cells( i + 1 + ii+j, 17 ) = Dev.GetComponentAttributeValue("Тех. описание 1")	' НЕ ЗАПОЛНЯЕТСЯ (9) Обозначение модуля / Module TAG
'			objExcel.Cells( i + 1 + ii+j, 18 ) = Dev.GetComponentAttributeValue("Тех. описание 1")	' НЕ ЗАПОЛНЯЕТСЯ (10) № канала / Channel №

' ===============================================================
' ===============================================================
' Подсчитываем количество заполненых строчек 
' ===============================================================
' ===============================================================
'i - строка
i = 0
ii= 5 'начинаем с 6-ой строки
j = 0
For i = 0 To 2000
	A2 = objExcel.Cells( i + 1 + ii, 2 )
		If A2 <> "" Then
			j = j + 1
		Else
			Exit for
		End If
Next 



kDI = 0
kDO = 0
kAI = 0
kAO = 0

kDI_0 = 0
kDO_0 = 0
kAI_0 = 0
kAO_0 = 0


' ===============================================================
' Подсчет количества каждого типа сигнала
' ===============================================================
For i = 0 To j - 1
	A11 = objExcel.Cells( i + 1 + ii, 8 )
	If A11 = "DI" Then 
		kDI = kDI + 1
	End If
	
	If A11 = "DO" Then 
		kDO = kDO + 1
	End if

	If A11 = "AI" Then 
		kAI = kAI + 1
	End if

	If A11 = "AO" Then 
		kAO = kAO + 1
	End if

	If A11 = "DI" Or A11 = "DO" Or A11 = "AI" Or A11 = "AO" Then 
	Else 
		kPR = kPR + 1
	End If
Next 



' ===============================================================
' Создаем массив сигналов из EXCEL
' ===============================================================
Redim ArrDeviceIds_ExcelDI(kDI, 6)
Redim ArrDeviceIds_ExcelDO(kDO, 6)
Redim ArrDeviceIds_ExcelAI(kAI, 6)
Redim ArrDeviceIds_ExcelAO(kAO, 6)
Redim ArrDeviceIds_ExcelPR(kPR, 6)


For i = 0 To j-1
'	A1 = objExcel.Cells( i + 1 + ii, 1 )
	A2 = objExcel.Cells( i + 1 + ii, 2 )			' (1) Позиционное обозначение / Position TAG
'	A4 = objExcel.Cells( i + 1 + ii, 3 )			' (2) Полный тег оборудования
'	A6 = objExcel.Cells( i + 1 + ii, 4 )			' (3) Признак резервирования
'	A6 = objExcel.Cells( i + 1 + ii, 5 )			' (4) Наименование сигнала / TAG
	A6 = objExcel.Cells( i + 1 + ii, 6 )			' (5) Описание сигнала / TAG discription
'	A9 = objExcel.Cells( i + 1 + ii, 7 )			' (6)  Внешнее подключение/Field connection 
	A11 = objExcel.Cells( i + 1 + ii, 8 )			' (7) Тип сигнала / Signal type
	A12 = objExcel.Cells( i + 1 + ii, 9 )			' (8) Тип подключения / Connection type
	A13 = objExcel.Cells( i + 1 + ii, 10 )			' (9) Ед. Изм / Measurement type
'	A14 = objExcel.Cells( i + 1 + ii, 14 )
'	A17 = objExcel.Cells( i + 1 + ii, 17 )
'	A18 = objExcel.Cells( i + 1 + ii, 18 )
	
	
	If A11 = "DI" Then 
		ArrDeviceIds_ExcelDI(kDI_0, 0) = A2					' (1) Позиционное обозначение / Position TAG
		ArrDeviceIds_ExcelDI(kDI_0, 1) = A6 				' (5) Описание сигнала / TAG discription
		ArrDeviceIds_ExcelDI(kDI_0, 2) = A11 				' (7) Тип сигнала / Signal type
		ArrDeviceIds_ExcelDI(kDI_0, 3) = A12 				' (8) Тип подключения / Connection type
		ArrDeviceIds_ExcelDI(kDI_0, 4) = A13 				' (9) Ед. Изм / Measurement type
		ArrDeviceIds_ExcelDI(kDI_0, 5) = 0 					' если 0 - можно испоьзовать, 1 - уже использован
		kDI_0 = kDI_0 + 1
	End If 
	
	If A11 = "DO" Then 
		ArrDeviceIds_ExcelDO(kDO_0, 0) = A2					' (1) Позиционное обозначение / Position TAG
		ArrDeviceIds_ExcelDO(kDO_0, 1) = A6 				' (3) Описание сигнала / TAG discription
		ArrDeviceIds_ExcelDO(kDO_0, 2) = A11 				' (5) Тип сигнала / Signal type
		ArrDeviceIds_ExcelDO(kDO_0, 3) = A12 				' (6) Тип подключения / Connection type
		ArrDeviceIds_ExcelDO(kDO_0, 4) = A13 				' (7) Ед. Изм / Measurement type
		ArrDeviceIds_ExcelDO(kDO_0, 5) = 0 					' если 0 - можно испоьзовать, 1 - уже использован
		kDO_0 = kDO_0 + 1
	End If 
'	
	If A11 = "AI" Then 
		ArrDeviceIds_ExcelAI(kAI_0, 0) = A2					' (1) Позиционное обозначение / Position TAG
		ArrDeviceIds_ExcelAI(kAI_0, 1) = A6 				' (3) Описание сигнала / TAG discription
		ArrDeviceIds_ExcelAI(kAI_0, 2) = A11 				' (5) Тип сигнала / Signal type
		ArrDeviceIds_ExcelAI(kAI_0, 3) = A12 				' (6) Тип подключения / Connection type
		ArrDeviceIds_ExcelAI(kAI_0, 4) = A13 				' (7) Ед. Изм / Measurement type
		ArrDeviceIds_ExcelAI(kAI_0, 5) = 0 				' если 0 - можно испоьзовать, 1 - уже использован
		kAI_0 = kAI_0 + 1
	End If 
'	
	If A11 = "AO" Then 
		ArrDeviceIds_ExcelAO(kAO_0, 0) = A2					' (1) Позиционное обозначение / Position TAG
		ArrDeviceIds_ExcelAO(kAO_0, 1) = A6 				' (3) Описание сигнала / TAG discription
		ArrDeviceIds_ExcelAO(kAO_0, 2) = A11 				' (5) Тип сигнала / Signal type
		ArrDeviceIds_ExcelAO(kAO_0, 3) = A12 				' (6) Тип подключения / Connection type
		ArrDeviceIds_ExcelAO(kAO_0, 4) = A13 				' (7) Ед. Изм / Measurement type
		ArrDeviceIds_ExcelAO(kAO_0, 5) = 0 					' если 0 - можно испоьзовать, 1 - уже использован
		kAO_0 = kAO_0 + 1
	End If 
	
'	If A11 = "DI" Or A11 = "DO" Or A11 = "AI" Or A11 = "AO" Then 
'	Else 
'		ArrDeviceIds_ExcelPR(i + kPR_0, 0) = A2					' (1) Позиционное обозначение / Position TAG
'		ArrDeviceIds_ExcelPR(i + kPR_0, 1) = A6 				' (3) Описание сигнала / TAG discription
'		ArrDeviceIds_ExcelPR(i + kPR_0, 2) = A11 				' (5) Тип сигнала / Signal type
'		ArrDeviceIds_ExcelPR(i + kPR_0, 3) = A12 				' (6) Тип подключения / Connection type
'		ArrDeviceIds_ExcelPR(i + kPR_0, 4) = A13 				' (7) Ед. Изм / Measurement type
'		kPR_0 = kPR_0 + 1
'	End If 
Next 












'================================================================================================
' Находим все пины ВАРИАНТ 2 DI
'================================================================================================
namePozObozDI = "DI"
namepinDI = "DI"
k = 0
i = 1

deviceCount = job.GetAllDeviceIds( deviceIds )        'get selected devices in the project tree
'deviceCount = job.GetTreeSelectedAllDeviceIds( deviceIds )        'get selected devices in the project tree
If deviceCount > 0 Then 
	For deviceIndex = 1 To deviceCount
'	If i <= kAI_0 Then
		deviceId = device.SetId( deviceIds( deviceIndex ) )
		deviceName = device.GetName()
		deviceNam = device.GetAttributeValue( "Позиционное обозначение (Вариант надписи)" )
		If InStr(1, deviceNam, namePozObozDI, 1) Then
			result = device.GetPinIds( pinIds )
			If result = 0 Then
				app.PutInfo 0, "No pins found for device item " & deviceName & " ( " & deviceId & " )"
				Else
				app.PutInfo 0, result & " pins found for device item " & deviceName & " ( " & deviceId & " ) :"
					For pinIndex = 1 To result
						attributeName = "ПК (PLC) - Физический адрес"
						pinId = pin.SetId( pinIds( pinIndex ) )
						pinName = pin.GetName()
						pinName_Attr = pin.GetAttributeValue( attributeName )
'						If pinName_Attr => namepinAI Then
						If InStr(1, pinName_Attr, namepinDI, 1) Then
							app.PutInfo 0, "    " & pinName & " ( " & pinId & " )"
							result1 = pin.GetAttributeValue( "TAG Позиция" )
							If result1 = "HOLD" Then
								If i <= kDI_0 Then	
									For k = 0 To kDI
										AA_0_1 = ArrDeviceIds_ExcelDI(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
										If AA_0_1 = 0 Then 
										
											A2_1 = ArrDeviceIds_ExcelDI(k, 0)				' (1) Позиционное обозначение / Position TAG
											A6_1 = ArrDeviceIds_ExcelDI(k, 1)				' (3) Описание сигнала / TAG discription
											A11_1 = ArrDeviceIds_ExcelDI(k, 2)				' (5) Тип сигнала / Signal type
											A12_1 = ArrDeviceIds_ExcelDI(k, 3)				' (6) Тип подключения / Connection type
											A13_1 = ArrDeviceIds_ExcelDI(k, 4)				' (7) Ед. Изм / Measurement type
											AA_0_1 = ArrDeviceIds_ExcelDI(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
											ArrDeviceIds_ExcelDI(k, 5) = 1				' если 0 - можно испоьзовать, 1 - уже использован
										i = i + 1
										Exit for
										End If
									Next 
									
										C11 = pin.SetAttributeValue( "TAG Позиция" , A2_1 ) ' Устанавливаем  новый атрибут
										C21 = pin.SetAttributeValue( "TAG Описание", A6_1 ) ' Устанавливаем  новый атрибут
										C31 = pin.SetAttributeValue( "ПЛК - Тип сигнала", A11_1 ) ' Устанавливаем  новый атрибут
										C41 = pin.SetAttributeValue( "ПЛК - Тип подключения", A12_1 ) ' Устанавливаем  новый атрибут
										C51 = pin.SetAttributeValue( "ПЛК - Единица измерения", A13_1 ) ' Устанавливаем  новый атрибут
								End If
							End If
						End If
					Next
			End If
		End If
'	Exit for
'	End If
	Next
End If


'================================================================================================
' Находим все пины ВАРИАНТ 2 DO
'================================================================================================
namePozObozDO = "DO"
namepinDI = "DO"

k = 0
i = 1

deviceCount = job.GetAllDeviceIds( deviceIds )        'get selected devices in the project tree
'deviceCount = job.GetTreeSelectedAllDeviceIds( deviceIds )        'get selected devices in the project tree
If deviceCount > 0 Then 
	For deviceIndex = 1 To deviceCount
'	If i <= kAI_0 Then
		deviceId = device.SetId( deviceIds( deviceIndex ) )
		deviceName = device.GetName()
		deviceNam = device.GetAttributeValue( "Позиционное обозначение (Вариант надписи)" )
		If InStr(1, deviceNam, namePozObozDO, 1) Then
			result = device.GetPinIds( pinIds )
			If result = 0 Then
				app.PutInfo 0, "No pins found for device item " & deviceName & " ( " & deviceId & " )"
				Else
				app.PutInfo 0, result & " pins found for device item " & deviceName & " ( " & deviceId & " ) :"
					For pinIndex = 1 To result
						attributeName = "ПК (PLC) - Физический адрес"
						pinId = pin.SetId( pinIds( pinIndex ) )
						pinName = pin.GetName()
						pinName_Attr = pin.GetAttributeValue( attributeName )
'						If pinName_Attr => namepinAI Then
						If InStr(1, pinName_Attr, namepinDO, 1) Then
							app.PutInfo 0, "    " & pinName & " ( " & pinId & " )"
							result1 = pin.GetAttributeValue( "TAG Позиция" )
							If result1 = "HOLD" Then
								If i <= kDO_0 Then	
									For k = 0 To kDO
										AA_0_1 = ArrDeviceIds_ExcelDO(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
										If AA_0_1 = 0 Then 
										
											A2_1 = ArrDeviceIds_ExcelDO(k, 0)				' (1) Позиционное обозначение / Position TAG
											A6_1 = ArrDeviceIds_ExcelDO(k, 1)				' (3) Описание сигнала / TAG discription
											A11_1 = ArrDeviceIds_ExcelDO(k, 2)				' (5) Тип сигнала / Signal type
											A12_1 = ArrDeviceIds_ExcelDO(k, 3)				' (6) Тип подключения / Connection type
											A13_1 = ArrDeviceIds_ExcelDO(k, 4)				' (7) Ед. Изм / Measurement type
											AA_0_1 = ArrDeviceIds_ExcelDO(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
											ArrDeviceIds_ExcelDO(k, 5) = 1				' если 0 - можно испоьзовать, 1 - уже использован
										i = i + 1
										Exit for
										End If
									Next 
									
										C11 = pin.SetAttributeValue( "TAG Позиция" , A2_1 ) ' Устанавливаем  новый атрибут
										C21 = pin.SetAttributeValue( "TAG Описание", A6_1 ) ' Устанавливаем  новый атрибут
										C31 = pin.SetAttributeValue( "ПЛК - Тип сигнала", A11_1 ) ' Устанавливаем  новый атрибут
										C41 = pin.SetAttributeValue( "ПЛК - Тип подключения", A12_1 ) ' Устанавливаем  новый атрибут
										C51 = pin.SetAttributeValue( "ПЛК - Единица измерения", A13_1 ) ' Устанавливаем  новый атрибут
								End If
							End If
						End If
					Next
			End If
		End If
'	Exit for
'	End If
	Next
End If


'================================================================================================
' Находим все пины ВАРИАНТ 2 AI
'================================================================================================
'namePozObozAI = InputBox("Введите обозначение стартового модуля AI", "", "")
'namepinAI = InputBox("Введите обозначение стартового входа модуля AI", "", "")

namePozObozAI = "AI"
namepinAI = "AI"
k = 0
i = 1

deviceCount = job.GetAllDeviceIds( deviceIds )        'get selected devices in the project tree
'deviceCount = job.GetTreeSelectedAllDeviceIds( deviceIds )        'get selected devices in the project tree
If deviceCount > 0 Then 
	For deviceIndex = 1 To deviceCount
'	If i <= kAI_0 Then
		deviceId = device.SetId( deviceIds( deviceIndex ) )
		deviceName = device.GetName()
		deviceNam = device.GetAttributeValue( "Позиционное обозначение (Вариант надписи)" )
		If InStr(1, deviceNam, namePozObozAI, 1) Then
			result = device.GetPinIds( pinIds )
			If result = 0 Then
				app.PutInfo 0, "No pins found for device item " & deviceName & " ( " & deviceId & " )"
				Else
				app.PutInfo 0, result & " pins found for device item " & deviceName & " ( " & deviceId & " ) :"
					For pinIndex = 1 To result
						attributeName = "ПК (PLC) - Физический адрес"
						pinId = pin.SetId( pinIds( pinIndex ) )
						pinName = pin.GetName()
						pinName_Attr = pin.GetAttributeValue( attributeName )
'						If pinName_Attr => namepinAI Then
						If InStr(1, pinName_Attr, namepinAI, 1) Then
							app.PutInfo 0, "    " & pinName & " ( " & pinId & " )"
							result1 = pin.GetAttributeValue( "TAG Позиция" )
							If result1 = "HOLD" Then
								If i <= kAI_0 Then	
									For k = 0 To kAI
										AA_0_1 = ArrDeviceIds_ExcelAI(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
										If AA_0_1 = 0 Then 
										
											A2_1 = ArrDeviceIds_ExcelAI(k, 0)				' (1) Позиционное обозначение / Position TAG
											A6_1 = ArrDeviceIds_ExcelAI(k, 1)				' (3) Описание сигнала / TAG discription
											A11_1 = ArrDeviceIds_ExcelAI(k, 2)				' (5) Тип сигнала / Signal type
											A12_1 = ArrDeviceIds_ExcelAI(k, 3)				' (6) Тип подключения / Connection type
											A13_1 = ArrDeviceIds_ExcelAI(k, 4)				' (7) Ед. Изм / Measurement type
											AA_0_1 = ArrDeviceIds_ExcelAI(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
											ArrDeviceIds_ExcelAI(k, 5) = 1				' если 0 - можно испоьзовать, 1 - уже использован
										i = i + 1
										Exit for
										End If
									Next 
									
										C11 = pin.SetAttributeValue( "TAG Позиция" , A2_1 ) ' Устанавливаем  новый атрибут
										C21 = pin.SetAttributeValue( "TAG Описание", A6_1 ) ' Устанавливаем  новый атрибут
										C31 = pin.SetAttributeValue( "ПЛК - Тип сигнала", A11_1 ) ' Устанавливаем  новый атрибут
										C41 = pin.SetAttributeValue( "ПЛК - Тип подключения", A12_1 ) ' Устанавливаем  новый атрибут
										C51 = pin.SetAttributeValue( "ПЛК - Единица измерения", A13_1 ) ' Устанавливаем  новый атрибут
								End If
							End If
						End If
					Next
			End If
		End If
'	Exit for
'	End If
	Next
End If

'================================================================================================
' Находим все пины ВАРИАНТ 2 AO
'================================================================================================
namePozObozAO = "AO"
namepinAI = "AO"

k = 0
i = 1

deviceCount = job.GetAllDeviceIds( deviceIds )        'get selected devices in the project tree
'deviceCount = job.GetTreeSelectedAllDeviceIds( deviceIds )        'get selected devices in the project tree
If deviceCount > 0 Then 
	For deviceIndex = 1 To deviceCount
'	If i <= kAO_0 Then
		deviceId = device.SetId( deviceIds( deviceIndex ) )
		deviceName = device.GetName()
		deviceNam = device.GetAttributeValue( "Позиционное обозначение (Вариант надписи)" )
		If InStr(1, deviceNam, namePozObozAO, 1) Then
			result = device.GetPinIds( pinIds )
			If result = 0 Then
				app.PutInfo 0, "No pins found for device item " & deviceName & " ( " & deviceId & " )"
				Else
				app.PutInfo 0, result & " pins found for device item " & deviceName & " ( " & deviceId & " ) :"
					For pinIndex = 1 To result
						attributeName = "ПК (PLC) - Физический адрес"
						pinId = pin.SetId( pinIds( pinIndex ) )
						pinName = pin.GetName()
						pinName_Attr = pin.GetAttributeValue( attributeName )
						If InStr(1, pinName_Attr, namepinAO, 1) Then
							app.PutInfo 0, "    " & pinName & " ( " & pinId & " )"
							result1 = pin.GetAttributeValue( "TAG Описание" )
							If result1 = "HOLD" Then
								If i <= kAO_0 Then
									For k = 0 To kAO
										AA_0_1 = ArrDeviceIds_ExcelAO(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
										If AA_0_1 = 0 Then 
											A2_1 = ArrDeviceIds_ExcelAO(k, 0)				' (1) Позиционное обозначение / Position TAG
											A6_1 = ArrDeviceIds_ExcelAO(k, 1)				' (3) Описание сигнала / TAG discription
											A11_1 = ArrDeviceIds_ExcelAO(k, 2)				' (5) Тип сигнала / Signal type
											A12_1 = ArrDeviceIds_ExcelAO(k, 3)				' (6) Тип подключения / Connection type
											A13_1 = ArrDeviceIds_ExcelAO(k, 4)				' (7) Ед. Изм / Measurement type
											AA_0_1 = ArrDeviceIds_ExcelAO(k, 5)				' если 0 - можно испоьзовать, 1 - уже использован
											ArrDeviceIds_ExcelAo(k, 5) = 1				' если 0 - можно испоьзовать, 1 - уже использован
										i = i + 1
										Exit for
										End If
									Next 
									C11 = pin.SetAttributeValue( "TAG Позиция" , A2_1 ) ' Устанавливаем  новый атрибут
									C21 = pin.SetAttributeValue( "TAG Описание", A6_1 ) ' Устанавливаем  новый атрибут
									C31 = pin.SetAttributeValue( "ПЛК - Тип сигнала", A11_1 ) ' Устанавливаем  новый атрибут
									C41 = pin.SetAttributeValue( "ПЛК - Тип подключения", A12_1 ) ' Устанавливаем  новый атрибут
									C51 = pin.SetAttributeValue( "ПЛК - Единица измерения", A13_1 ) ' Устанавливаем  новый атрибут
								
								End If
							End If
						End If
					Next
			End If
		End If
'	Exit for
'	End If
	Next
End If





'	wscript.Quit






'====================================================================================================================
App.PutMessage "=================================================" 
App.PutMessage "Генерация выполнена успешно!" 
App.PutMessage "=================================================" 

ExitScript true, job




Set signal = Nothing
Set Dev2 = Nothing
Set Dev1 = Nothing
Set Pin2 = Nothing
Set Pin1 = Nothing
Set Sheet = Nothing
Set Pin = Nothing
Set Cor = Nothing
Set Cab = Nothing

Set tree = Nothing
Set graphic = Nothing
Set sheet = Nothing
Set sym = Nothing
Set symbol = Nothing
Set dev = Nothing
Set device = Nothing
Set job = Nothing
Set app = Nothing

	wscript.Quit






