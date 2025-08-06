' ===============================================================

'Скрипт предназначен для переименования клеммников и клемм согласно отчета F20, согласно типа модуля.
'Скрипт работает со следующими модулями:
'- R500 AI 08 052-000-AAA
'- R500 AI 16 012-000-AAA
'- R500 AO 08 011-000-AAA

' ===============================================================




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
' Для EXCEL
' ===============================================================

'Dim ExcelApp
'	Dim Excel, ExcelName
'Dim xlThin, xlThick, xlDouble, xlAutomatic, xlLandscape, xlCenter, xlLeft, xlRight, HorizontalAlignment

'xlCenter = -4108
'xlLeft = -4131

'Set Excel = CreateObject("Excel.Application") 	' создаем объект Excel
'Excel.Visible = 1								' делаем Excel видимым 
'Excel.WorkBooks.Add								' создаем новую Книгу в Excel
'Excel.Cells( 1, 1 ) = "Имя жилы"
'Excel.Cells( 1, 2 ) = "Марка провода"
'Excel.Cells( 1, 3 ) = "Поз.обоз 1"
'Excel.Cells( 1, 4 ) = "Вывод 1"
'Excel.Cells( 1, 5 ) = "Поз.обоз 2"
'Excel.Cells( 1, 6 ) = "Вывод 2"
'Excel.Cells( 1, 7 ) = "Имя цепи"
'Excel.Cells( 1, 8 ) = "1С"
'Excel.Cells( 1, 9 ) = "ПК (PLC)"
'Excel.Cells( 1, 10 ) = "ПК (PLC)"
'Excel.Cells( 1, 11 ) = "Поз.обоз 3"
'Excel.Cells( 1, 12 ) = "Вывод 3"
' ===============================================================
' ===============================================================




' ===============================================================
' Находим все изделия типа "МОДУЛЬ КОНТРОЛЛЕРА AI или AO"
' ===============================================================
app.PutInfo 0, "==============================="
app.PutInfo 0, "Модули AO. ШАГ 1 из 4"
app.PutInfo 0, "==============================="
k_modAI8 = 0
k_modAI16 = 0
k_modAO8 = 0
Redim ArrDeviceIds_1(100, 4)
deviceCnt = job.GetAllDeviceIds(deviceIds)
For deviceIndex = 1 To deviceCnt
	deviceId = device.SetId( deviceIds( deviceIndex ) )
	deviceName = device.GetName()
	result1 = device.GetComponentName()
	deviceNam = device.GetAttributeValue( "Позиционное обозначение (Вариант надписи)" )
'	If InStr(1, deviceNam, "AI", 1) Then
'	symbolFilter = 0
'	result2 = device.GetSymbolIds( symbolIds, symbolFilter )
'	app.PutInfo 0, "Device " & deviceName & " ( " & deviceId & " ) has " & result & " symbols: "
'		For symbolIndex = 1 To result2        'loop through all found symbols
'		symbolId = symbol.SetId( symbolIds( symbolIndex ) )
'		symbolName = symbol.GetName()
'		
'		result3 = symbol.GetSchemaLocation( xPosition, yPosition, gridDescription, column, row )
'		result4 = symbol.GetSymbolTypeName()
'		sheet.SetId result3
'		result5 = sheet.GetName
'			If result4 = "Модуль_аналогового_ввода_ R500_AI_08_052-000-AAA" Then 
'				app.PutInfo 0, "Изделие: " & result1 & "; Символ: "  & symbolName & ", " & result4 & "( " & symbolId & " ): " & " Страница: " & sheet.GetName
'				ArrDeviceIds_1(k, 0) = deviceId				' ID изделия
'				ArrDeviceIds_1(k, 1) = deviceName			' Наименование в на схеме
'				ArrDeviceIds_1(k, 2) = result4				' Наименование символа
'				ArrDeviceIds_1(k, 3) = result5				' Номер страницы
'				k = k + 1
'				k_modAI8 = k_modAI8 + 1
'			End If 
'			If result4 = "R500_AI_16_012-000-AAA_1" Or result4 = "R500_AI_16_012-000-AAA_2" Then 
'				app.PutInfo 0, "Изделие: " & result1 & "; Символ: "  & symbolName & ", " & result4 & "( " & symbolId & " ): " & " Страница: " & sheet.GetName
'				ArrDeviceIds_1(k, 0) = deviceId				' ID изделия
'				ArrDeviceIds_1(k, 1) = deviceName			' Наименование в на схеме
'				ArrDeviceIds_1(k, 2) = result4				' Наименование символа
'				ArrDeviceIds_1(k, 3) = result5				' Номер страницы
'				k = k + 1
'				k_modAI16 = k_modAI16 + 1
'			End If 
'			
'		Next
'	End If 
	
	If InStr(1, deviceNam, "AO", 1) Then
	symbolFilter = 0
	result2 = device.GetSymbolIds( symbolIds, symbolFilter )
	app.PutInfo 0, "Device " & deviceName & " ( " & deviceId & " ) has " & result & " symbols: "
		For symbolIndex = 1 To result2        'loop through all found symbols
		symbolId = symbol.SetId( symbolIds( symbolIndex ) )
		symbolName = symbol.GetName()
		
		result3 = symbol.GetSchemaLocation( xPosition, yPosition, gridDescription, column, row )
		result4 = symbol.GetSymbolTypeName()
		sheet.SetId result3
		result5 = sheet.GetName
			If result4 = "Модуль_аналогового_вывода_R500_AO_08_011-000-AAA" Then 
				app.PutInfo 0, "Изделие: " & result1 & "; Символ: "  & symbolName & ", " & result4 & "( " & symbolId & " ): " & " Страница: " & sheet.GetName
				ArrDeviceIds_1(k, 0) = deviceId				' ID изделия
				ArrDeviceIds_1(k, 1) = deviceName			' Наименование в на схеме
				ArrDeviceIds_1(k, 2) = result4				' Наименование символа
				ArrDeviceIds_1(k, 3) = result5				' Номер страницы
				k = k + 1
				k_modAO8 = k_modAO8 + 1
			End If 
		Next
	End If 
Next

' ===============================================================
' Определяем общее количество аналоговых входов и выходов
' ===============================================================
kanalAI8 = k_modAI8 * 8
kanalAI16 = k_modAI16 * 16
kanalAO8 = k_modAO8 * 8

kanalItog = kanalAI8 + kanalAI16 + kanalAO8

k_Arr_1 = k
'k_Arr_1 = 0
'For k=0 To 100
'	If ArrDeviceIds_1(k, 3) <> "" Then 
'		k_Arr_1 = k_Arr_1+1
'	End if
'Next	


'	wscript.Quit

' ===============================================================
' Находим все соединения
' ===============================================================
app.PutInfo 0, "Модули AO. ШАГ 2 из 4"
app.PutInfo 0, "==============================="

k=0
Redim ArrDeviceIds_10(1000, 19)
CabCnt = Job.GetCableIds (cabIds)
For s = 0 To k_Arr_1-1
	SHEET_n = ArrDeviceIds_1(s, 3)

'App.ClearOutputWindow
'App.PutMessage "Начало скрипта " 

'CabCnt = .GetCableIds (cabIds)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
For i=1 to CabCnt
	Cab.SetId cabIds(i)
	PinCnt=Cab.GetPinIds(pinIds)
	For j=1 to PinCnt
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
		Pin.SetId pinIds(j)
		Pin1.SetId Pin.GetEndPinId( 1 )
		Dev1.SetId Pin.GetEndPinId( 1 )
		Pin2.SetId Pin.GetEndPinId( 2 )
		Dev2.SetId Pin.GetEndPinId( 2 )
		

		result100 = pin1.GetSchemaLocation( xPosition, yPosition, gridDescription, column, row )
		sheet.SetId result100
		result2 = pin.GetSignalName()
					A0 = Cab.GetName				' "Имя жилы провода" - "Провода" или Поз. обоз. кабеля
					A1 = Pin.GetName				' "Марка провода" - Наименование на схеме
					A2 = Dev1.GetName				' "Поз.обоз 1" - Откуда
					A3 = Pin1.GetName				' "Поз.обоз 1" - Вывод1 - Откуда
					A4 = Dev1.GetAttributeValue( "Позиционное обозначение (Вариант надписи)" )
					A5 = Pin1.GetAttributeValue( "ПК (PLC) - Физический адрес" )
					A6 = Dev2.GetName				' "Поз.обоз 2" - куда
					A7 = Pin2.GetName				' "Поз.обоз 2" - Вывод2 - куда
					A8 = result2					' "Имя цепи"
					A9 = Dev1.GetComponentName()	' Наименование изделия 1 в Е3
					A10 = sheet.GetName
					
					A11 = Pin1.GetAttributeValue( "ПК (PLC) - Физический адрес" )
					A12 = Pin1.GetAttributeValue( "TAG Позиция" )
					A13 = Pin1.GetAttributeValue( "TAG Описание" )
					A14 = Pin1.GetAttributeValue( "ПЛК - Тип сигнала" )

'					A15 = Pin2.GetAttributeValue( "ПК (PLC) - Физический адрес" )
'					A16 = Pin2.GetAttributeValue( "TAG Позиция" )
'					A17 = Pin2.GetAttributeValue( "TAG Описание" )
'					A18 = Pin2.GetAttributeValue( "ПЛК - Тип сигнала" )


				If SHEET_n = A10 Then 
						ArrDeviceIds_10(k, 0) = A0				' "Имя жилы провода" - "Провода" или Поз. обоз. кабеля
						ArrDeviceIds_10(k, 1) = A1 				' "Марка провода" - Наименование на схеме
						ArrDeviceIds_10(k, 2) = A2 				' "Поз.обоз 1" - Откуда
						ArrDeviceIds_10(k, 3) = A3 				' "Поз.обоз 1" - Вывод1 - Откуда
						ArrDeviceIds_10(k, 4) = A4 				' "Позиционное обозначение (Вариант надписи)" для модулей контроллера
						ArrDeviceIds_10(k, 5) = A5 				' "ПК (PLC) - Физический адрес"
						ArrDeviceIds_10(k, 6) = A6 				' "Поз.обоз 2" - куда
						ArrDeviceIds_10(k, 7) = A7 				' "Поз.обоз 2" - Вывод2 - куда
						ArrDeviceIds_10(k, 8) = A8 				' "Имя цепи"
						ArrDeviceIds_10(k, 9) = A9 				' Наименование изделия 1 в Е3
						ArrDeviceIds_10(k, 10) = A10 			' Номер страницы
						ArrDeviceIds_10(k, 11) = 0 				' 0 или 1,Использование
						
						ArrDeviceIds_10(k, 12) = A11 			' PIN 1 - "ПК (PLC) - Физический адрес" 
						ArrDeviceIds_10(k, 13) = A12 			' PIN 1 - "TAG Позиция"
						ArrDeviceIds_10(k, 14) = A13 			' PIN 1 - "TAG Описание"
						ArrDeviceIds_10(k, 15) = A14 			' PIN 1 - "ПЛК - Тип сигнала"
						
'						ArrDeviceIds_10(k, 16) = A15 			' PIN 2 - "ПК (PLC) - Физический адрес" 
'						ArrDeviceIds_10(k, 17) = A16 			' PIN 2 - "TAG Позиция"
'						ArrDeviceIds_10(k, 18) = A17 			' PIN 2 - "TAG Описание"
'						ArrDeviceIds_10(k, 19) = A18 			' PIN 2 - "ПЛК - Тип сигнала"
						
						k = k +1
				End  If 
	Next
Next
Next

k_Arr_10 = k
'k_Arr_10 = 0
'For k=0 To 1000
'	If ArrDeviceIds_10(k, 2) <> "" Then 
'		k_Arr_10 = k_Arr_10 + 1
'	End if
'Next

'wscript.Quit


' ===============================================================
' Исключаем промtжуточные соединения и клеммы
' ===============================================================
app.PutInfo 0, "Модули AO. ШАГ 3 из 4"
app.PutInfo 0, "==============================="
Redim ArrDeviceIds_11(kanalItog-1, 17)
l = 0

For k = 0 To k_Arr_10-1

	A0 = ArrDeviceIds_10(k, 0)				' "Имя жилы провода" - "Провода" или Поз. обоз. кабеля
	A1 = ArrDeviceIds_10(k, 1)				' "Марка провода" - Наименование на схеме
	A2 = ArrDeviceIds_10(k, 2)				' "Поз.обоз 1" - Откуда
	A3 = ArrDeviceIds_10(k, 3)				' "Поз.обоз 1" - Вывод1 - Откуда
	A4 = ArrDeviceIds_10(k, 4)				' "Позиционное обозначение (Вариант надписи)" для модулей контроллера
	A5 = ArrDeviceIds_10(k, 5)				' "ПК (PLC) - Физический адрес"
	A6 = ArrDeviceIds_10(k, 6)				' "Поз.обоз 2" - куда
	A7 = ArrDeviceIds_10(k, 7)				' "Поз.обоз 2" - Вывод2 - куда
	A8 = ArrDeviceIds_10(k, 8)				' "Имя цепи"
	A9 = ArrDeviceIds_10(k, 9)				' Наименование изделия 1 в Е3
	A10 = ArrDeviceIds_10(k, 10)			' Номер страницы
	A11 = ArrDeviceIds_10(k, 11)				' 0 или 1,Использование

	A12 = ArrDeviceIds_10(k, 12)			' PIN 1 - "ПК (PLC) - Физический адрес" 
	A13 = ArrDeviceIds_10(k, 13)			' PIN 1 - "TAG Позиция"
	A14 = ArrDeviceIds_10(k, 14)			' PIN 1 - "TAG Описание"
	A15 = ArrDeviceIds_10(k, 15)			' PIN 1 - "ПЛК - Тип сигнала"
	
'	A16 = ArrDeviceIds_10(k, 16)			' PIN 2 - "ПК (PLC) - Физический адрес" 
'	A17 = ArrDeviceIds_10(k, 17)			' PIN 2 - "TAG Позиция"
'	A18 = ArrDeviceIds_10(k, 18)			' PIN 2 - "TAG Описание"
'	A19 = ArrDeviceIds_10(k, 19)			' PIN 2 - "ПЛК - Тип сигнала"


'	If A4 = "AI" Then 
'		If InStr(1, A12, "AI", 1) Then
'			For k1= 0 To k_Arr_10-1
'			A2_1 = ArrDeviceIds_10(k1, 2)				' "Поз.обоз 1" - Откуда
'			A3_1 = ArrDeviceIds_10(k1, 3)				' "Поз.обоз 1" - Вывод1 - Откуда
'			A6_1 = ArrDeviceIds_10(k1, 6)				' "Поз.обоз 2" - Куда
'			A7_1 = ArrDeviceIds_10(k1, 7)				' "Поз.обоз 2" - Вывод2 - куда
'			A8_1 = ArrDeviceIds_10(k1, 8)				' "Имя цепи"
'		
'				If A6 = A2_1 And A7 = A3_1 Then
'					ArrDeviceIds_11(l, 0) = A0				' "Имя жилы провода" - "Провода" или Поз. обоз. кабеля
'					ArrDeviceIds_11(l, 1) = A1				' "Марка провода" - Наименование на схеме
'					ArrDeviceIds_11(l, 2) = A2				' "Поз.обоз 1" - Откуда
'					ArrDeviceIds_11(l, 3) = A3				' "Поз.обоз 1" - Вывод1 - Откуда
'					ArrDeviceIds_11(l, 4) = A4				' "Позиционное обозначение (Вариант надписи)" для модулей контроллера
'					ArrDeviceIds_11(l, 6) = A6_1			' "Поз.обоз 2" - куда
'					ArrDeviceIds_11(l, 7) = A7_1			' "Поз.обоз 2" - Вывод2 - куда
'					ArrDeviceIds_11(l, 8) = A8				' "Имя цепи"
'			
'					ArrDeviceIds_11(l, 9) = A9				' Наименование изделия 1 в Е3
'					ArrDeviceIds_11(l, 10) = A10			' Номер страницы
'					ArrDeviceIds_11(l, 11) = 0				' 0 или 1,Использование
'					
'					ArrDeviceIds_11(l, 12) = A12 			' PIN 1 - "ПК (PLC) - Физический адрес" 
'					ArrDeviceIds_11(l, 13) = A13 			' PIN 1 - "TAG Позиция"
'					ArrDeviceIds_11(l, 14) = A14 			' PIN 1 - "TAG Описание"
'					ArrDeviceIds_11(l, 15) = A15 			' PIN 1 - "ПЛК - Тип сигнала"
'					l = l + 1
'					
'					ArrDeviceIds_10(k, 11) = 1 				' 0 или 1,Использование
'					ArrDeviceIds_10(k1, 11) = 1 			' 0 или 1,Использование
'				End If
'			Next
'		End If
'	End If


	If A4 = "AO" Then 
		If InStr(1, A12, "AO", 1) Then
			For k1= 0 To k_Arr_10-1
			A2_1 = ArrDeviceIds_10(k1, 2)				' "Поз.обоз 1" - Откуда
			A3_1 = ArrDeviceIds_10(k1, 3)				' "Поз.обоз 1" - Вывод1 - Откуда
			A6_1 = ArrDeviceIds_10(k1, 6)				' "Поз.обоз 2" - Куда
			A7_1 = ArrDeviceIds_10(k1, 7)				' "Поз.обоз 2" - Вывод2 - куда
			A8_1 = ArrDeviceIds_10(k1, 8)				' "Имя цепи"
		
				If A6 = A2_1 And A7 = A3_1 Then
					ArrDeviceIds_11(l, 0) = A0				' "Имя жилы провода" - "Провода" или Поз. обоз. кабеля
					ArrDeviceIds_11(l, 1) = A1				' "Марка провода" - Наименование на схеме
					ArrDeviceIds_11(l, 2) = A2				' "Поз.обоз 1" - Откуда
					ArrDeviceIds_11(l, 3) = A3				' "Поз.обоз 1" - Вывод1 - Откуда
					ArrDeviceIds_11(l, 4) = A4				' "Позиционное обозначение (Вариант надписи)" для модулей контроллера
					ArrDeviceIds_11(l, 6) = A6_1			' "Поз.обоз 2" - куда
					ArrDeviceIds_11(l, 7) = A7_1			' "Поз.обоз 2" - Вывод2 - куда
					ArrDeviceIds_11(l, 8) = A8				' "Имя цепи"
			
					ArrDeviceIds_11(l, 9) = A9				' Наименование изделия 1 в Е3
					ArrDeviceIds_11(l, 10) = A10			' Номер страницы
					ArrDeviceIds_11(l, 11) = 0				' 0 или 1,Использование
					
					ArrDeviceIds_11(l, 12) = A12 			' PIN 1 - "ПК (PLC) - Физический адрес" 
					ArrDeviceIds_11(l, 13) = A13 			' PIN 1 - "TAG Позиция"
					ArrDeviceIds_11(l, 14) = A14 			' PIN 1 - "TAG Описание"
					ArrDeviceIds_11(l, 15) = A15 			' PIN 1 - "ПЛК - Тип сигнала"
					l = l + 1
					
					ArrDeviceIds_10(k, 11) = 1 				' 0 или 1,Использование
					ArrDeviceIds_10(k1, 11) = 1 			' 0 или 1,Использование
				End If
			Next
		End If
	End If


Next








'app.PutInfo 0, "Модули AI, AO. ШАГ 3 из 8"
'wscript.Quit



' ===============================================================
' Находим выходной клеммник
' ===============================================================

'Redim ArrDeviceKL_1(2000, 4)
m = 0

Ndevs = Job.GetTerminalIds ( DevIds )				' получаем массив ID всех клеммников (DevIds) в проекте и их кол-во (Ndevs) 
Redim ArrDeviceKL_1(Ndevs, 4)
For n = 1 To Ndevs 
		Dev.Setid DevIds(n)
		KL1 = Dev.GetName 											'Поз. обозначение клеммника
		KL2 = Dev.GetMasterPinName									' Номер клеммы
		KL3 = Dev.GetComponentName									' Обозначение в БД Е3
		KL4 = Dev.GetComponentAttributeValue("Наименование")		' Обозначение в 1С
		KL5 = Dev.GetComponentAttributeValue("Тех. описание 1")		'Сечение
		If KL1 <> "" And KL2 <> "" Then 
			ArrDeviceKL_1(m, 0) = KL1
			ArrDeviceKL_1(m, 1) = KL2
			ArrDeviceKL_1(m, 2) = KL3
			ArrDeviceKL_1(m, 3) = KL4
			ArrDeviceKL_1(m, 4) = KL5
			m = m + 1
		End If
		
Next 

k_KL_1 = m

' ===============================================================
' Дополняем массив ArrDeviceIds_13 клемамми, которые находятся перед клеммами выходов контроллера
' ===============================================================
For l = 0 To kanalItog-1
	A0 = ArrDeviceIds_11(l, 0)				' "Имя жилы провода" - "Провода" или Поз. обоз. кабеля
	A1 = ArrDeviceIds_11(l, 1)				' "Марка провода" - Наименование на схеме
	A2 = ArrDeviceIds_11(l, 2)				' "Поз.обоз 1" - Откуда
	A3 = ArrDeviceIds_11(l, 3)				' "Поз.обоз 1" - Вывод1 - Откуда
	A4 = ArrDeviceIds_11(l, 4)				' "Позиционное обозначение (Вариант надписи)" для модулей контроллера
	A6 = ArrDeviceIds_11(l, 6)				' "Поз.обоз 2" - куда
	A7 = ArrDeviceIds_11(l, 7)				' "Поз.обоз 2" - Вывод2 - куда
	A8 = ArrDeviceIds_11(l, 8)				' "Имя цепи"
	
	A9 = ArrDeviceIds_11(l, 9)				' Наименование изделия 1 в Е3
	A10 = ArrDeviceIds_11(l, 10)			' Номер страницы
	A11 = ArrDeviceIds_11(l, 11)			' 0 или 1,Использование
	
'	A12 = ArrDeviceIds_11(l, 12)			' PIN 1 - "ПК (PLC) - Физический адрес" 
'	A13 = ArrDeviceIds_11(l, 13)			' PIN 1 - "TAG Позиция"
'	A14 = ArrDeviceIds_11(l, 14)			' PIN 1 - "TAG Описание"
'	A15 = ArrDeviceIds_11(l, 15)			' PIN 1 - "ПЛК - Тип сигнала"
	
		For j = 0 To k_KL_1
			KL1 = ArrDeviceKL_1(j, 0) 				'Поз. обозначение клеммника
			KL2 = ArrDeviceKL_1(j, 1)				' Номер клеммы
			KL3 = ArrDeviceKL_1(j, 2)				' Обозначение в БД Е3
			KL4 = ArrDeviceKL_1(j, 3)				' Обозначение в 1С
			KL5 = ArrDeviceKL_1(j, 4)				'Сечение
			
			KL1_1 = ArrDeviceKL_1(j+1, 0) 				'Поз. обозначение клеммника
			KL2_1 = ArrDeviceKL_1(j+1, 1)				' Номер клеммы
	
	
'			If A6 = KL1 And A7 = KL2 Then
'				If A9 = "Модуль_8AI_R500_AI_08_052-000-AAA" Or A9 = "Модуль_16AI_R500_AI_16_011-000-AAA" Or A9 = "Модуль_16AI_R500_AI_16_012-000-AAA" Then
'					ArrDeviceIds_11(l, 16) = ArrDeviceKL_1(j-1, 0)				' "Поз.обоз 2" - Куда - Добавочный клеммник
'					ArrDeviceIds_11(l, 17) = ArrDeviceKL_1(j-1, 1)				' "Поз.обоз 2" - Куда - Добавочный клеммник - клемма
'					Exit For 
'				End If 
'			End If 
			
			If A6 = KL1 And A7 = KL2 Then
				If A9 = "Модуль_8AO_R500_AO_08_011-000-AAA" Then
					ArrDeviceIds_11(l, 16) = ArrDeviceKL_1(j+1, 0)				' "Поз.обоз 2" - Куда - Добавочный клеммник
					ArrDeviceIds_11(l, 17) = ArrDeviceKL_1(j+1, 1)				' "Поз.обоз 2" - Куда - Добавочный клеммник - клемма
					Exit For 
				End If 
			End If 
		Next
Next



'wscript.Quit



' ===============================================================
' Для EXCEL
' ===============================================================
'For m = 0 To k_kol_3-1
'		
'		Excel.Cells( m + 2, 1 ) = ArrDeviceIds_13(m, 0) 											'Поз. обозначение клеммника
'		Excel.Cells( m + 2, 2 ) = ArrDeviceIds_13(m, 1)									' Номер клеммы
'		Excel.Cells( m + 2, 3 ) = ArrDeviceIds_13(m, 2)									' Обозначение в БД Е3
'		Excel.Cells( m + 2, 4 ) = ArrDeviceIds_13(m, 3)		' Обозначение в 1С
'		Excel.Cells( m + 2, 5 ) = ArrDeviceIds_13(m, 4)	'Сечение
'		Excel.Cells( m + 2, 6 ) = ArrDeviceIds_13(m, 5)
'		Excel.Cells( m + 2, 7 ) = ArrDeviceIds_13(m, 6)
'		Excel.Cells( m + 2, 8 ) = ArrDeviceIds_13(m, 7)
'		Excel.Cells( m + 2, 9 ) = ArrDeviceIds_13(m, 8)
'		Excel.Cells( m + 2, 10 ) = ArrDeviceIds_13(m, 9)
'		Excel.Cells( m + 2, 11 ) = ArrDeviceIds_13(m, 10)
'		Excel.Cells( m + 2, 12 ) = ArrDeviceIds_13(m, 11)
'Next 

'Set Excel = Nothing
' ===============================================================
' ===============================================================


'--------------------------------------------------------------------------------------------
app.PutInfo 0, "Модули AO. ШАГ 4 из 4"
app.PutInfo 0, "==============================="

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


' ===============================================================
' Подсчитываем количество заполненых строчек 
' ===============================================================
' ===============================================================
i = 0
ii= 5
j = 0
For i = 0 To 1000
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
'	If A11 = "DI" Then 
'		kDI = kDI + 1
'	End If
	
'	If A11 = "DO" Then 
'		kDO = kDO + 1
'	End if

'	If A11 = "AI" Then 
'		kAI = kAI + 1
'	End if

	If A11 = "AO" Then 
		kAO = kAO + 1
	End if

'	If A11 = "DI" Or A11 = "DO" Or A11 = "AI" Or A11 = "AO" Then 
'	Else 
'		kPR = kPR + 1
'	End If
Next 



' ===============================================================
' Создаем массив сигналов из EXCEL
' ===============================================================
'Redim ArrDeviceIds_ExcelDI(kDI, 4)
'Redim ArrDeviceIds_ExcelDO(kDO, 4)
Redim ArrDeviceIds_ExcelAI(kAI, 4)
Redim ArrDeviceIds_ExcelAO(kAO, 4)
'Redim ArrDeviceIds_ExcelPR(kPR, 4)



For i = 0 To j-1
	A1 = objExcel.Cells( i + 1 + ii, 2 )					' (1) Позиционное обозначение / Position TAG
	A2 = objExcel.Cells( i + 1 + ii, 6 )					' (5) Описание сигнала / TAG discription
	A3 = objExcel.Cells( i + 1 + ii, 7 )					' (6) Внешнее подключение / Field connection
	A4 = objExcel.Cells( i + 1 + ii, 8 )					' (7) Тип сигнала / Signal type

'	If A11 = "DI" Then 
'		ArrDeviceIds_ExcelDI(kDI_0, 0) = A1					' (1) Позиционное обозначение / Position TAG
'		ArrDeviceIds_ExcelDI(kDI_0, 1) = A2 				' (5) Описание сигнала / TAG discription
'		ArrDeviceIds_ExcelDI(kDI_0, 2) = A3 				' (6) Внешнее подключение / Field connection
'		ArrDeviceIds_ExcelDI(kDI_0, 3) = A4 				' (7) Тип сигнала / Signal type
'		ArrDeviceIds_ExcelDI(kDI_0, 4) = 0 					' если 0 - можно испоьзовать, 1 - уже использован
'		kDI_0 = kDI_0 + 1
'	End If 
'	
'	If A11 = "DO" Then 
'		ArrDeviceIds_ExcelDO(kDO_0, 0) = A1					' (1) Позиционное обозначение / Position TAG
'		ArrDeviceIds_ExcelDO(kDO_0, 1) = A2 				' (3) Описание сигнала / TAG discription
'		ArrDeviceIds_ExcelDO(kDO_0, 2) = A3 				' (4) Внешнее подключение / Field connection
'		ArrDeviceIds_ExcelDO(kDO_0, 3) = A4 				' (5) Тип сигнала / Signal type
'		ArrDeviceIds_ExcelDO(kDO_0, 4) = 0 					' если 0 - можно испоьзовать, 1 - уже использован
'		kDO_0 = kDO_0 + 1
'	End If 
'	
'	If A4 = "AI" Then 
'		ArrDeviceIds_ExcelAI(kAI_0, 0) = A1					' (1) Позиционное обозначение / Position TAG
'		ArrDeviceIds_ExcelAI(kAI_0, 1) = A2 				' (5) Описание сигнала / TAG discription
'		ArrDeviceIds_ExcelAI(kAI_0, 2) = A3 				' (6) Внешнее подключение / Field connection
'		ArrDeviceIds_ExcelAI(kAI_0, 3) = A4 				' (7) Тип сигнала / Signal type
'		ArrDeviceIds_ExcelAI(kAI_0, 4) = 0 					' если 0 - можно испоьзовать, 1 - уже использован
'		kAI_0 = kAI_0 + 1
'	End If 
'	
	If A4 = "AO" Then 
		ArrDeviceIds_ExcelAO(kAO_0, 0) = A1					' (1) Позиционное обозначение / Position TAG
		ArrDeviceIds_ExcelAO(kAO_0, 1) = A2 				' (5) Описание сигнала / TAG discription
		ArrDeviceIds_ExcelAO(kAO_0, 2) = A3 				' (6) Внешнее подключение / Field connection
		ArrDeviceIds_ExcelAO(kAO_0, 3) = A4 				' (7) Тип сигнала / Signal type
		ArrDeviceIds_ExcelAO(kAO_0, 4) = 0 					' если 0 - можно испоьзовать, 1 - уже использован
		kAO_0 = kAO_0 + 1
	End If 
	
Next 

'================================================================================================
' Находим все пины ВАРИАНТ 3 AI
'================================================================================================
'For i = 0 To kAI_0 - 1
'		D1 = ArrDeviceIds_ExcelAI(i, 0)					' (1) Позиционное обозначение / Position TAG
'		D2 = ArrDeviceIds_ExcelAI(i, 1)					' (3) Описание сигнала / TAG discription
'		D3 = ArrDeviceIds_ExcelAI(i, 2)					' (4) Внешнее подключение / Field connection
'		D4 = ArrDeviceIds_ExcelAI(i, 3)					' (5) Тип сигнала / Signal type
'		D5 = ArrDeviceIds_ExcelAI(i, 4)					' если 0 - можно испоьзовать, 1 - уже использован
'		
'		For j = 0 To kanalItog-1
'		E9 = ArrDeviceIds_11(j, 12)						' "ПК (PLC) - Физический адрес" 
'		E10 = ArrDeviceIds_11(j, 13)						' "TAG Позиция"
'		E11 = ArrDeviceIds_11(j, 14)					' "TAG Описание"
'		E12 = ArrDeviceIds_11(j, 15)					' "ПЛК - Тип сигнала"
'		
'		E5 = ArrDeviceIds_11(j, 6)						'ПЕРВЫЙ КЛЕММНИК + КЛЕММА
'		E6 = ArrDeviceIds_11(j, 7)						'КЛЕММА первого клеммника
'		E13 = ArrDeviceIds_11(j, 16)					'ВТОРОЙ КЛЕММНИК + КЛЕММА
'		E14 = ArrDeviceIds_11(j, 17)					'КЛЕММА второго клеммника
'		E15 = ArrDeviceIds_11(j, 11)					' если 0 - можно испоьзовать, 1 - уже использован
'		
'		If D5 = 0 And E15 = 0 Then 
'		If D1 = E10 And D2 = E11 And D4 = E12 Then 
'		A3_1 = D3
'		A3_5_1 = ""
'		A3_5_2 = ""
'		A3_7_1 = ""
'		A3_7_2 = ""
'		A3_9_1 = ""
'		A3_9_2 = ""
'		
'			A3_1_1 = Len(A3_1)												' Количество символов в строке
'			A3_1_2 = InStr(1, A3_1, ";", vbTextCompare)						' Номер символа в строке ";", начиная с первого - это первый символ ";"
'			If A3_1_2 > 0 Then
'				A3_2_2 = InStr(A3_1_2 + 1, A3_1,";", vbTextCompare)			' Находим второй символ в строке ";", начиная с "A3_1_2 + 1"
'				If A3_2_2 > 0 Then
'					 Если три клеммы в F20
'					A3_3 = A3_2_2 - A3_1_2									' Длина второго клеммника
'					
'					A3_4_1 = Left(A3_1, A3_1_2 - 1)							' Получаем все символы до ";" - ПЕРВЫЙ КЛЕММНИК + КЛЕММА
'					A3_5 = InStr(1, A3_4_1, ":", vbTextCompare)
'					A3_5_0 = Len(A3_4_1)
'					A3_5_1 = Left(A3_4_1, A3_5 - 1)						'!! 1 !! Получаем все символы до ":" - ПЕРВЫЙ КЛЕММНИК
'					A3_5_2 = Right(A3_4_1, A3_5_0 - A3_5)				'!! 1 !! Получаем все символы до ":" - КЛЕММА первого клеммника
'					
'					A3_6_1 = Mid(A3_1, A3_1_2 + 2, A3_3 - 2)					' Получаем все символы между ";" - ВТОРОЙ КЛЕММНИК + КЛЕММА
'					A3_7 = InStr(1, A3_6_1, ":", vbTextCompare)
'					A3_7_0 = Len(A3_6_1)
'					A3_7_1 = Left(A3_6_1, A3_7 - 1)						'!! 2 !! Получаем все символы до ":" - ВТОРОЙ КЛЕММНИК
'					A3_7_2 = Right(A3_6_1, A3_7_0 - A3_7)				'!! 2 !! Получаем все символы до ":" - КЛЕММА второго клеммника

'					A3_8_1 = Right(A3_1, A3_1_1 - A3_2_2)						' Получаем все символы между ";" - ТРЕТИЙ КЛЕММНИК + КЛЕММА
'					A3_9 = InStr(1, A3_8_1, ":", vbTextCompare)
'					A3_9_0 = Len(A3_8_1)
'					A3_9_1 = Left(A3_8_1, A3_9 - 1)						'!! 3 !! Получаем все символы до ":" - ТРЕТИЙ КЛЕММНИК
'					A3_9_2 = Right(A3_8_1, A3_9_0 - A3_9)				'!! 3 !! Получаем все символы до ":" - КЛЕММА третьего клеммника

'				Else 
'					 Если две клеммы в F20
'					A3_4_1 = Left(A3_1, A3_1_2 - 1)							' Получаем все символы до ";" - ПЕРВЫЙ КЛЕММНИК
'					A3_5 = InStr(1, A3_4_1, ":", vbTextCompare)
'					A3_5_0 = Len(A3_4_1)
'					A3_5_1 = Left(A3_4_1, A3_5 - 1)						'!! 1 !! Получаем все символы до ":" - ПЕРВЫЙ КЛЕММНИК
'					A3_5_2 = Right(A3_4_1, A3_5_0 - A3_5)				'!! 1 !! Получаем все символы до ":" - КЛЕММА первого клеммника
'															
'					A3_4_2 = Right(A3_1, A3_1_1 - A3_1_2-1)						' Получаем все символы между ";" - ВТОРОЙ КЛЕММНИК
'					A3_7 = InStr(1, A3_4_2, ":", vbTextCompare)
'					A3_7_0 = Len(A3_4_2)
'					A3_7_1 = Left(A3_4_2, A3_7 - 1)						'!! 2 !! Получаем все символы до ":" - ВТОРОЙ КЛЕММНИК
'					A3_7_2 = Right(A3_4_2, A3_7_0 - A3_7)				'!! 2 !! Получаем все символы до ":" - КЛЕММА второго клеммника
'				End If 
'												
'				Else 
'				 Если одна клеммы в F20
'					A3_5 = InStr(1, A3_1, ":", vbTextCompare)
'					A3_5_0 = Len(A3_1)
'					A3_5_1 = Left(A3_1, A3_5 - 1)						'!! 1 !! Получаем все символы до ":" - ПЕРВЫЙ КЛЕММНИК
'					A3_5_2 = Right(A3_1, A3_5_0 - A3_5)					'!! 1 !! Получаем все символы до ":" - КЛЕММА первого клеммника
'			End If

'				Ndevs = Job.GetTerminalIds ( DevIds )				' получаем массив ID всех клеммников (DevIds) в проекте и их кол-во (Ndevs) 
'				For n = 1 To Ndevs 
'					Dev.Setid DevIds(n)
'					KL10 = Dev.GetName 											'Поз. обозначение клеммника
'					KL20 = Dev.GetMasterPinName									' Номер клеммы
'					If InStr(1, KL10, "XT", 1) Then

'						If E5 <> "" And E6 <> "" And E13 <> "" And E14 <> "" Then 
'							If A3_5_1 <> "" And A3_5_2 <> "" And A3_7_1 <> "" And A3_7_2 <> ""Then 
'								If KL10 = E5 And KL20 = E6 Then 
'									result11 = dev.SetName( A3_7_1 )
'									result12 = dev.SetMasterPinName( A3_7_2 )
'									ArrDeviceIds_ExcelAI(i, 4) = 1
'									ArrDeviceIds_11(j, 11) = 1
'								End If 
'		
'								If KL10 = E13 And KL20 = E14 Then 
'									result13 = dev.SetName( A3_5_1 )
'									result14 = dev.SetMasterPinName( A3_5_2 )
'									ArrDeviceIds_ExcelAI(i, 4) = 1
'									ArrDeviceIds_11(j, 11) = 1
'								End If 
'							End If 
'					
'							If A3_5_1 <> "" And A3_5_2 <> "" And A3_7_1 = "" And A3_7_2 = ""Then 
'								If KL10 = E5 And KL20 = E6 Then 
'									result11 = dev.SetName( A3_5_1 )
'									result12 = dev.SetMasterPinName( A3_5_2 )
'									ArrDeviceIds_ExcelAI(i, 4) = 1
'									ArrDeviceIds_11(j, 11) = 1
'								End If 
'								If KL10 = E13 And KL20 = E14 Then 
'								A3_7_1 = ""
'								A3_7_2 = ""
'									result11 = dev.SetName( A3_7_1 )
'									result12 = dev.SetMasterPinName( A3_7_2 )
'									ArrDeviceIds_ExcelAI(i, 4) = 1
'									ArrDeviceIds_11(j, 11) = 1
'								End If 
'							End If 
'						End If 
'					End If 
'				Next 
'		End If
'		End If
'	Next
'Next


'================================================================================================
' Находим все пины ВАРИАНТ 3 AO
'================================================================================================
For i = 0 To kAO_0 - 1
		D1 = ArrDeviceIds_ExcelAO(i, 0)					' (1) Позиционное обозначение / Position TAG
		D2 = ArrDeviceIds_ExcelAO(i, 1)					' (3) Описание сигнала / TAG discription
		D3 = ArrDeviceIds_ExcelAO(i, 2)					' (4) Внешнее подключение / Field connection
		D4 = ArrDeviceIds_ExcelAO(i, 3)					' (5) Тип сигнала / Signal type
		D5 = ArrDeviceIds_ExcelAO(i, 4)					' если 0 - можно испоьзовать, 1 - уже использован
		
		For j = 0 To kanalItog-1
		E9 = ArrDeviceIds_11(j, 12)						' "ПК (PLC) - Физический адрес" 
		E10 = ArrDeviceIds_11(J, 13)						' "TAG Позиция"
		E11 = ArrDeviceIds_11(j, 14)					' "TAG Описание"
		E12 = ArrDeviceIds_11(j, 15)					' "ПЛК - Тип сигнала"
		
		E5 = ArrDeviceIds_11(j, 6)						'ПЕРВЫЙ КЛЕММНИК + КЛЕММА
		E6 = ArrDeviceIds_11(j, 7)						'КЛЕММА первого клеммника
		E13 = ArrDeviceIds_11(j, 16)					'ВТОРОЙ КЛЕММНИК + КЛЕММА
		E14 = ArrDeviceIds_11(j, 17)					'КЛЕММА второго клеммника
		E15 = ArrDeviceIds_11(j, 11)					' если 0 - можно испоьзовать, 1 - уже использован

		If D5 = 0 And E15 = 0 Then 
		If D1 = E10 And D2 = E11 And D4 = E12 Then 
		A3_1 = D3
		A3_5_1 = ""
		A3_5_2 = ""
		A3_7_1 = ""
		A3_7_2 = ""
		A3_9_1 = ""
		A3_9_2 = ""
		
			A3_1_1 = Len(A3_1)												' Количество символов в строке
			A3_1_2 = InStr(1, A3_1, ";", vbTextCompare)						' Номер символа в строке ";", начиная с первого - это первый символ ";"
			If A3_1_2 > 0 Then
				A3_2_2 = InStr(A3_1_2 + 1, A3_1,";", vbTextCompare)			' Находим второй символ в строке ";", начиная с "A3_1_2 + 1"
				If A3_2_2 > 0 Then
					' Если три клеммы в F20
					A3_3 = A3_2_2 - A3_1_2									' Длина второго клеммника
					
					A3_4_1 = Left(A3_1, A3_1_2 - 1)							' Получаем все символы до ";" - ПЕРВЫЙ КЛЕММНИК + КЛЕММА
					A3_5 = InStr(1, A3_4_1, ":", vbTextCompare)
					A3_5_0 = Len(A3_4_1)
					A3_5_1 = Left(A3_4_1, A3_5 - 1)						' Получаем все символы до ":" - ПЕРВЫЙ КЛЕММНИК
					A3_5_2 = Right(A3_4_1, A3_5_0 - A3_5)				' Получаем все символы до ":" - КЛЕММА первого клеммника
					
					A3_6_1 = Mid(A3_1, A3_1_2 + 2, A3_3 - 2)					' Получаем все символы между ";" - ВТОРОЙ КЛЕММНИК + КЛЕММА
					A3_7 = InStr(1, A3_6_1, ":", vbTextCompare)
					A3_7_0 = Len(A3_6_1)
					A3_7_1 = Left(A3_6_1, A3_7 - 1)						' Получаем все символы до ":" - ВТОРОЙ КЛЕММНИК
					A3_7_2 = Right(A3_6_1, A3_7_0 - A3_7)				' Получаем все символы до ":" - КЛЕММА второго клеммника

					A3_8_1 = Right(A3_1, A3_1_1 - A3_2_2)						' Получаем все символы между ";" - ТРЕТИЙ КЛЕММНИК + КЛЕММА
					A3_9 = InStr(1, A3_8_1, ":", vbTextCompare)
					A3_9_0 = Len(A3_8_1)
					A3_9_1 = Left(A3_8_1, A3_9 - 1)						' Получаем все символы до ":" - ТРЕТИЙ КЛЕММНИК
					A3_9_2 = Right(A3_8_1, A3_9_0 - A3_9)				' Получаем все символы до ":" - КЛЕММА третьего клеммника

				Else 
					' Если две клеммы в F20
					A3_4_1 = Left(A3_1, A3_1_2 - 1)							' Получаем все символы до ";" - ПЕРВЫЙ КЛЕММНИК
					A3_5 = InStr(1, A3_4_1, ":", vbTextCompare)
					A3_5_0 = Len(A3_4_1)
					A3_5_1 = Left(A3_4_1, A3_5 - 1)						' Получаем все символы до ":" - ПЕРВЫЙ КЛЕММНИК
					A3_5_2 = Right(A3_4_1, A3_5_0 - A3_5)					' Получаем все символы до ":" - КЛЕММА первого клеммника
															
					A3_4_2 = Right(A3_1, A3_1_1 - A3_1_2-1)						' Получаем все символы между ";" - ВТОРОЙ КЛЕММНИК
					A3_7 = InStr(1, A3_4_2, ":", vbTextCompare)
					A3_7_0 = Len(A3_4_2)
					A3_7_1 = Left(A3_4_2, A3_7 - 1)						' Получаем все символы до ":" - ВТОРОЙ КЛЕММНИК
					A3_7_2 = Right(A3_4_2, A3_7_0 - A3_7)					' Получаем все символы до ":" - КЛЕММА второго клеммника
				End If 
												
				Else 
				' Если одна клеммы в F20
					A3_5 = InStr(1, A3_1, ":", vbTextCompare)
					A3_5_0 = Len(A3_1)
					A3_5_1 = Left(A3_1, A3_5 - 1)						' Получаем все символы до ":" - ПЕРВЫЙ КЛЕММНИК
					A3_5_2 = Right(A3_1, A3_5_0 - A3_5)					' Получаем все символы до ":" - КЛЕММА первого клеммника
			End If


				Ndevs = Job.GetTerminalIds ( DevIds )				' получаем массив ID всех клеммников (DevIds) в проекте и их кол-во (Ndevs) 
				For n = 1 To Ndevs 
					Dev.Setid DevIds(n)
					KL10 = Dev.GetName 											'Поз. обозначение клеммника
					KL20 = Dev.GetMasterPinName									' Номер клеммы
					If InStr(1, KL10, "XT", 1) Then

						If E5 <> "" And E6 <> "" And E13 <> "" And E14 <> "" Then 
							If A3_5_1 <> "" And A3_5_2 <> "" And A3_7_1 <> "" And A3_7_2 <> ""Then 
								If KL10 = E5 And KL20 = E6 Then 
									result11 = dev.SetName( A3_5_1 )
									result12 = dev.SetMasterPinName( A3_5_2 )
									ArrDeviceIds_ExcelAO(i, 4) = 1
									ArrDeviceIds_11(j, 11) = 1
								End If 
		
								If KL10 = E13 And KL20 = E14 Then 
									result13 = dev.SetName( A3_7_1 )
									result14 = dev.SetMasterPinName( A3_7_2 )
									ArrDeviceIds_ExcelAO(i, 4) = 1
									ArrDeviceIds_11(j, 11) = 1
								End If 
							End If 
					
							If A3_5_1 <> "" And A3_5_2 <> "" And A3_7_1 = "" And A3_7_2 = ""Then 
								If KL10 = E5 And KL20 = E6 Then 
									result11 = dev.SetName( A3_5_1 )
									result12 = dev.SetMasterPinName( A3_5_2 )
									ArrDeviceIds_ExcelAO(i, 4) = 1
									ArrDeviceIds_11(j, 11) = 1
								End If 
'								If KL10 = E13 And KL20 = E14 Then 
'								A3_7_1 = ""
'								A3_7_2 = ""
'									result13 = dev.SetName( A3_7_1 )
'									result14 = dev.SetMasterPinName( A3_7_2 )
'									ArrDeviceIds_ExcelAO(i, 4) = 1
'									ArrDeviceIds_11(j, 11) = 1
'								End If 
'								Exit For 
							End If 
						End If 
					End If 
				Next 
		End If
		End If
	Next
Next















'====================================================================================================================
'App.PutMessage "=================================================" 
App.PutMessage "Генерация выполнена успешно!" 
'App.PutMessage "=================================================" 

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


