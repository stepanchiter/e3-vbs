' ===============================================================
' СПЕЦПРОЕКТ, генерация схем REGUL и компоновка
' ===============================================================


' ===============================================================
' Создание глобальных переменных
Dim app, appId

' Функция подключения к E3
'Set app = CreateObject( "CT.Application" ) 
Set app = E3Connection()
Set job = app.CreateJobObject()
Set sheet = job.CreateSheetObject()

Set dev = job.CreateDeviceObject
Set device = job.CreateDeviceObject()
Set subDevice = job.CreateDeviceObject()
Set pin = job.CreatePinObject
Set connection = job.CreateConnectionObject()
Set devicePin = job.CreatePinObject()

Set slot = job.CreateSlotObject()
Set symbol = job.CreateSymbolObject()
Set component = job.CreateComponentObject()
Set sym = job.CreateSymbolObject()







' Процедура выхода
'Call ExitScript(True, 0)


' ===============================================================
' Функция подключения к E3
' ===============================================================
Function E3Connection()
'	 Улавливатель ошибок
	On Error Resume Next
'		 Подключение процессу E3
		Set app = CreateObject("CT.Application")
'		 Запрос идентификатора процесса
		appId = app.GetId()
'		 Вывод сообщения об ошибке
		If (appId = 0) Then
'			 Вывод сообщения
			MsgBox "Ошибка. Процесс E3.series не запущен или COM-классы приложения E3.series не зарегистрированы!", 16, "Ошибка"
		End If
	On Error Goto 0
'	 Возврат функции
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
'Sub ExitScript(ByVal flagSuccessExit, ByRef job) 
	' Проверка флага
'	If (flagSuccessExit) Then
		' Вывод сообщения об успешном окончании
'		Call app.PutInfo(0, "Автоматизация выполнена успешно!")
'	Else
		' Вывод сообщения о не успешном окончании
'		Call app.PutError(1, "Автоматизация не выполнена!")
'	End If
'	
	' Очистка объектов
'	Set job = Nothing
'	Set app = Nothing
	' Выход из скрипта
'	WScript.Quit
'End Sub

' ===============================================================
' Запрос у пользователя: Обозначение документа
' ===============================================================
'OBOZNACHENIE = InputBox("Обозначение документа", "", "")
' If OBOZNACHENIE = "" Then 
' 	app.PutInfo 0, " ==============================================================="
'	app.PutError 0, OBOZNACHENIE
'	app.PutError 0, "Ошибка: Обозначение документа. Работа остановлена" 
'	app.PutError 0, " ==============================================================="
'	wscript.Quit
' wscript.Quit
' End if

' ===============================================================
' Создание проекта с именем файла "OBOZNACHENIE"
' ===============================================================
'Call Include("ВЕРСА_200_Создание файла.vbs") 


' ===============================================================
' Выполнение файла проверки маркировки
' ===============================================================
'Call Include("SPEC_Markirovka_proverka_v1.vbs") 




Set objExcel = CreateObject("Excel.Application")

' ===============================================================
' ===============================================================
' Открытие файла "______.xlsx"
' В запросе надо указать путь файла
' ===============================================================
' ===============================================================

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
	Dim thisFolder: thisFolder = fso.GetParentFolderName(WScript.ScriptFullName)
	' Полный путь до нужного файлами
	Dim fileFullName
	fileFullName = InputBox("Введите путь к файлу EXCEL с модулями контроллера", "", "")
'	 fileFullName = "D:\БХК\Script\Распределение контроллера.xlsx"
'	fileFullName = fso.BuildPath(thisFolder, fileName)
	' Проверка файла
	If (fso.FileExists(fileFullName)) Then
'		objExcel.Visible = False 
'		objExcel.Visible = True 
		'fileFullName = "D:\БХК\Script\Распределение контроллера.xlsx"
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
objExcel.Worksheets("Start").Activate

' ===============================================================
' Определяем тип соединения для модулей - провод или кабель
' ===============================================================
SHIFR_soed = InputBox("Введите тип соединения для модулей контроллера:" & vbNewLine & "1 - соединения проводом" & vbNewLine & "2 - соединения кабелем", "", "")


' ===============================================================
' Определяем количество модулей по типам
' ===============================================================


Kol_AI_kanal = objExcel.Cells( 3, 3 ) ' Общее количество каналов
Kol_AQ_kanal = objExcel.Cells( 4, 3 ) ' Общее количество каналов
Kol_DI_kanal = objExcel.Cells( 5, 3 ) ' Общее количество каналов
Kol_DQ_kanal = objExcel.Cells( 6, 3 ) ' Общее количество каналов

Kol_AI_mod = objExcel.Cells( 3, 4 ) ' Общее количество модулей
Kol_AQ_mod = objExcel.Cells( 4, 4 ) ' Общее количество модулей
Kol_DI_mod = objExcel.Cells( 5, 4 ) ' Общее количество модулей
Kol_DQ_mod = objExcel.Cells( 6, 4 ) ' Общее количество модулей



' ===============================================================
' Определение количества листов
' ===============================================================
Kol_sheet = 0
Kol_sh_AI_mod = 0
Kol_sh_AQ_mod = 0
Kol_sh_DI_mod = 0
Kol_sh_DQ_mod = 0

Kol_sh_AI_mod = Kol_sh_AI_mod * 1
Kol_sh_AQ_mod = Kol_sh_AQ_mod * 1
Kol_sh_DI_mod = Kol_sh_DI_mod * 4
Kol_sh_DQ_mod = Kol_sh_DQ_mod * 4

Kol_sheet = Kol_sh_AI_mod * 1 + Kol_sh_AQ_mod * 1 + Kol_sh_DI_mod * 4 + Kol_sh_DQ_mod * 4

' ===============================================================
' ВСТАВКА СТРАНИЦ ДЛЯ Э3
' ===============================================================
'Call Vstavka_stranits
'wscript.Quit

' ===============================================================
' Создаем массив модулей для КРЕЙТА 1
' ===============================================================

kol1 = objExcel.Cells(3, 8 )' Здесь надо взять ссылку с листа - Количество модулей
Redim ArrDeviceIds_ExcelKR1(kol1, 11)
kol11 = kol1
kol1 = 0
For i = 0 To kol11 - 1

KR_1_shassi = objExcel.Cells( i + 21, 3 )				' Шасси №п/п
KR_1_shassiCatalog = objExcel.Cells( i + 21, 4 )		' Шасси № каталог
KR_1_modul = objExcel.Cells( i + 21, 5 )				' Модуль 
KR_1_modulCatalog = objExcel.Cells( i + 21, 6 )			' Модуль № каталог
KR_1_kolodka = objExcel.Cells( i + 21, 7 )				' Колодка
KR_1_kolodkaCatalog = objExcel.Cells( i + 21, 8 )		' Колодка № каталог
KR_1_rele = objExcel.Cells( i + 21, 9 )					' Реле
KR_1_klemmnik1 = objExcel.Cells( i + 21, 10 )			' Клеммник 1
KR_1_klemmnik2 = objExcel.Cells( i + 21, 11 )			' Клеммник 2
KR_1_kabel1 = objExcel.Cells( i + 21, 12 )				' Кабель 1
KR_1_kabel2 = objExcel.Cells( i + 21, 13 )				' Кабель 2

	ArrDeviceIds_ExcelKR1(kol1, 0) = KR_1_shassi				' Поз.обозн. Шасси
	ArrDeviceIds_ExcelKR1(kol1, 1) = KR_1_shassiCatalog			' Поз.обозн. Шасси № каталог
	ArrDeviceIds_ExcelKR1(kol1, 2) = KR_1_modul 				' Поз.обозн. Модуль
	ArrDeviceIds_ExcelKR1(kol1, 3) = KR_1_modulCatalog 			' Поз.обозн. Модуль № каталог
	ArrDeviceIds_ExcelKR1(kol1, 4) = KR_1_kolodka 				' Поз.обозн. Колодка
	ArrDeviceIds_ExcelKR1(kol1, 5) = KR_1_kolodkaCatalog 		' Поз.обозн. Колодка № каталог
	ArrDeviceIds_ExcelKR1(kol1, 6) = KR_1_rele 					' Поз.обозн. Реле
	ArrDeviceIds_ExcelKR1(kol1, 7) = KR_1_klemmnik1 			' Поз.обозн. Клеммник 1
	ArrDeviceIds_ExcelKR1(kol1, 8) = KR_1_klemmnik2 			' Поз.обозн. Клеммник 2
	ArrDeviceIds_ExcelKR1(kol1, 9) = KR_1_kabel1 				' Поз.обозн. Кабель 1
	ArrDeviceIds_ExcelKR1(kol1, 10) = KR_1_kabel2 				' Поз.обозн. Кабель 2
	kol1 = kol1 + 1
Next

' ===============================================================
' Создаем массив модулей для КРЕЙТА 2
' ===============================================================

kol2 = objExcel.Cells(4, 8 )' Здесь надо взять ссылку с листа - Количество модулей
Redim ArrDeviceIds_ExcelKR2(kol2, 11)
kol21 = kol2
kol2 = 0
For i = 0 To kol21 - 1

KR_2_shassi = objExcel.Cells( i + 41, 3 )				' Шасси №п/п
KR_2_shassiCatalog = objExcel.Cells( i + 41, 4 )		' Шасси № каталог
KR_2_modul = objExcel.Cells( i + 41, 5 )				' Модуль 
KR_2_modulCatalog = objExcel.Cells( i + 41, 6 )			' Модуль № каталог
KR_2_kolodka = objExcel.Cells( i + 41, 7 )				' Колодка
KR_2_kolodkaCatalog = objExcel.Cells( i + 41, 8 )		' Колодка № каталог
KR_2_rele = objExcel.Cells( i + 41, 9 )					' Реле
KR_2_klemmnik1 = objExcel.Cells( i + 41, 10 )			' Клеммник 1
KR_2_klemmnik2 = objExcel.Cells( i + 41, 11 )			' Клеммник 2
KR_2_kabel1 = objExcel.Cells( i + 41, 12 )				' Кабель 1
KR_2_kabel2 = objExcel.Cells( i + 41, 13 )				' Кабель 2

	ArrDeviceIds_ExcelKR2(kol2, 0) = KR_2_shassi				' Поз.обозн. Шасси
	ArrDeviceIds_ExcelKR2(kol2, 1) = KR_2_shassiCatalog			' Поз.обозн. Шасси № каталог
	ArrDeviceIds_ExcelKR2(kol2, 2) = KR_2_modul 				' Поз.обозн. Модуль
	ArrDeviceIds_ExcelKR2(kol2, 3) = KR_2_modulCatalog 			' Поз.обозн. Модуль № каталог
	ArrDeviceIds_ExcelKR2(kol2, 4) = KR_2_kolodka 				' Поз.обозн. Колодка
	ArrDeviceIds_ExcelKR2(kol2, 5) = KR_2_kolodkaCatalog 		' Поз.обозн. Колодка № каталог
	ArrDeviceIds_ExcelKR2(kol2, 6) = KR_2_rele 					' Поз.обозн. Реле
	ArrDeviceIds_ExcelKR2(kol2, 7) = KR_2_klemmnik1 			' Поз.обозн. Клеммник 1
	ArrDeviceIds_ExcelKR2(kol2, 8) = KR_2_klemmnik2 			' Поз.обозн. Клеммник 2
	ArrDeviceIds_ExcelKR2(kol2, 9) = KR_2_kabel1 				' Поз.обозн. Кабель 1
	ArrDeviceIds_ExcelKR2(kol2, 10) = KR_2_kabel2 				' Поз.обозн. Кабель 2
	kol2 = kol2 + 1
Next

' ===============================================================
' Создаем массив модулей для КРЕЙТА 3
' ===============================================================

kol3 = objExcel.Cells(5, 8 )' Здесь надо взять ссылку с листа - Количество модулей
Redim ArrDeviceIds_ExcelKR3(kol3, 11)
kol31 = kol3
kol3 = 0
For i = 0 To kol31 - 1

KR_3_shassi = objExcel.Cells( i + 61, 3 )				' Шасси №п/п
KR_3_shassiCatalog = objExcel.Cells( i + 61, 4 )		' Шасси № каталог
KR_3_modul = objExcel.Cells( i + 61, 5 )				' Модуль 
KR_3_modulCatalog = objExcel.Cells( i + 61, 6 )			' Модуль № каталог
KR_3_kolodka = objExcel.Cells( i + 61, 7 )				' Колодка
KR_3_kolodkaCatalog = objExcel.Cells( i + 61, 8 )		' Колодка № каталог
KR_3_rele = objExcel.Cells( i + 61, 9 )					' Реле
KR_3_klemmnik1 = objExcel.Cells( i + 61, 10 )			' Клеммник 1
KR_3_klemmnik2 = objExcel.Cells( i + 61, 11 )			' Клеммник 2
KR_3_kabel1 = objExcel.Cells( i + 61, 12 )				' Кабель 1
KR_3_kabel2 = objExcel.Cells( i + 61, 13 )				' Кабель 2

	ArrDeviceIds_ExcelKR3(kol3, 0) = KR_3_shassi				' Поз.обозн. Шасси
	ArrDeviceIds_ExcelKR3(kol3, 1) = KR_3_shassiCatalog			' Поз.обозн. Шасси № каталог
	ArrDeviceIds_ExcelKR3(kol3, 2) = KR_3_modul 				' Поз.обозн. Модуль
	ArrDeviceIds_ExcelKR3(kol3, 3) = KR_3_modulCatalog 			' Поз.обозн. Модуль № каталог
	ArrDeviceIds_ExcelKR3(kol3, 4) = KR_3_kolodka 				' Поз.обозн. Колодка
	ArrDeviceIds_ExcelKR3(kol3, 5) = KR_3_kolodkaCatalog 		' Поз.обозн. Колодка № каталог
	ArrDeviceIds_ExcelKR3(kol3, 6) = KR_3_rele 					' Поз.обозн. Реле
	ArrDeviceIds_ExcelKR3(kol3, 7) = KR_3_klemmnik1 			' Поз.обозн. Клеммник 1
	ArrDeviceIds_ExcelKR3(kol3, 8) = KR_3_klemmnik2 			' Поз.обозн. Клеммник 2
	ArrDeviceIds_ExcelKR3(kol3, 9) = KR_3_kabel1 				' Поз.обозн. Кабель 1
	ArrDeviceIds_ExcelKR3(kol3, 10) = KR_3_kabel2 				' Поз.обозн. Кабель 2
	kol3 = kol3 + 1
Next


' ===============================================================
' ===============================================================
' Вставка модулей КРЕЙТА 1 на лист
' ===============================================================
kol1 = objExcel.Cells(3, 8 ) ' Определяем количество модулей в крейте 1
If kol1 > 0 Then 
i = 0
For i = 0 To kol1
BB0 = ArrDeviceIds_ExcelKR1(i, 0) 			' Шасси №п/п
BB1 = ArrDeviceIds_ExcelKR1(i, 1)			' Шасси № каталог
BB2 = ArrDeviceIds_ExcelKR1(i, 2)			' Модуль 
BB3 = ArrDeviceIds_ExcelKR1(i, 3)			' Модуль № каталог
BB4 = ArrDeviceIds_ExcelKR1(i, 4)			' Колодка
BB5 = ArrDeviceIds_ExcelKR1(i, 5)			' Колодка № каталог
BB6 = ArrDeviceIds_ExcelKR1(i, 6)			' Реле
BB7 = ArrDeviceIds_ExcelKR1(i, 7)			' Клеммник 1
BB8 = ArrDeviceIds_ExcelKR1(i, 8)			' Клеммник 2
BB9 = ArrDeviceIds_ExcelKR1(i, 9)			' Кабель 1
BB10 = ArrDeviceIds_ExcelKR1(i, 10)			' Кабель 2


'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AI 08 052-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AI08_052-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AI08_052-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-1000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-1000AI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-1000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-1000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-1000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-1000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If
'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AI 16 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AI16_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AI16_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_AI16_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-1001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-1001AI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-1001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-1001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-1001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-1001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
	
	name = "-1001W2"
	deviceDesignation = BB10
	Call Pereimenovanie
	
End If


'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AO 08 011-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AO08_052-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AO08_052-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-2000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-2000AO1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-2000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-2000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-2000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-2000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If

'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DI 32 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DI32_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-3001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-3001DI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-3001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-3001KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-3001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-3001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-3001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If


'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DI 32 013-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DI32_013-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-3001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-3001DI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-3001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-3001KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-3001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-3001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-3001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If

'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DO 32 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DO32_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-4000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-4000DO1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-4000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-4000KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-4000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-4000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-4000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If
Next 
End If
' ======================================================================================================



Sub  Pereimenovanie
	deviceId = device.Search(name, assignment, location)
	deviceName = device.GetName()
	result = device.SetName( deviceDesignation )
End Sub

Sub  Pereimenovanie_controller
	deviceId = device.Search(name, assignment, location)
	deviceName = device.GetName()
	result = device.SetName( deviceDesignation )
End sub


Sub  Pereimenovanie_Chassi
Dim componentVersion : componentVersion = "1"
	deviceId = device.Search(name, assignment, location)
	deviceName = device.GetName()
	result2 = device.SetName( deviceDesignation ) ' Переименование Поз.Обозначения
	result2_1 = device.GetComponentName() ' Наименование в БД Е3
	
	objExcel.Worksheets("Perechen").Activate
	For j=1 To 120
		If BB1 = objExcel.Cells( j, 4 ) Then 
			AA = objExcel.Cells( j, 2 )
			result2_2 = device.SetComponentName( AA, componentVersion )
			Exit for
		End If
	Next
	
	objExcel.Worksheets("Start").Activate
	
End sub





Function Pereimenovanie_KL
deviceCnt = job.GetAllDeviceIds(deviceIds)
For deviceIndex = 1 To deviceCnt
	deviceId = device.SetId( deviceIds( deviceIndex ) )
	
	deviceName1 = device.GetName()
	
	If InStr(1, deviceName1, name, 1) Then
	AA = Replace (deviceName1, name, deviceDesignation, 1, -1)
	result = device.SetName( AA )

	End If

Next
End Function


'===============================================================================================


'wscript.Quit




















' ===============================================================
' ===============================================================
' Вставка модулей КРЕЙТА 2 на лист
' ===============================================================
kol2 = objExcel.Cells(4, 8 ) ' Определяем количество модулей в крейте 2
If kol2 > 0 Then 
i = 0
For i = 0 To kol2
BB0 = ArrDeviceIds_ExcelKR2(i, 0)
BB1 = ArrDeviceIds_ExcelKR2(i, 1)
BB2 = ArrDeviceIds_ExcelKR2(i, 2)
BB3 = ArrDeviceIds_ExcelKR2(i, 3)
BB4 = ArrDeviceIds_ExcelKR2(i, 4)
BB5 = ArrDeviceIds_ExcelKR2(i, 5)
BB6 = ArrDeviceIds_ExcelKR2(i, 6)
BB7 = ArrDeviceIds_ExcelKR2(i, 7)
BB8 = ArrDeviceIds_ExcelKR2(i, 8)
BB9 = ArrDeviceIds_ExcelKR2(i, 9)
BB10 = ArrDeviceIds_ExcelKR2(i, 10)

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AI 08 052-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AI08_052-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AI08_052-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-1000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-1000AI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-1000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-1000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-1000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-1000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If
'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AI 16 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AI16_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AI16_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_AI16_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-1001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-1001AI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-1001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-1001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-1001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-1001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
	
	name = "-1001W2"
	deviceDesignation = BB10
	Call Pereimenovanie
	
End If


'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AO 08 011-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AO08_052-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AO08_052-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-2000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-2000AO1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-2000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-2000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-2000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-2000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If

'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DI 32 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DI32_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-3001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-3001DI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-3001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-3001KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-3001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-3001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-3001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If


'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DI 32 013-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DI32_013-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-3001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-3001DI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-3001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-3001KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-3001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-3001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-3001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If

'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DO 32 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DO32_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-4000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-4000DO1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-4000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-4000KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-4000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-4000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-4000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If
Next 
End If
' ======================================================================================================


'wscript.Quit

' ===============================================================
' ===============================================================
' Вставка модулей КРЕЙТА 3 на лист
' ===============================================================
kol3 = objExcel.Cells(5, 8 ) ' Определяем количество модулей в крейте 2
If kol3 > 0 Then 
i = 0
For i = 0 To kol3
BB0 = ArrDeviceIds_ExcelKR3(i, 0)
BB1 = ArrDeviceIds_ExcelKR3(i, 1)
BB2 = ArrDeviceIds_ExcelKR3(i, 2)
BB3 = ArrDeviceIds_ExcelKR3(i, 3)
BB4 = ArrDeviceIds_ExcelKR3(i, 4)
BB5 = ArrDeviceIds_ExcelKR3(i, 5)
BB6 = ArrDeviceIds_ExcelKR3(i, 6)
BB7 = ArrDeviceIds_ExcelKR3(i, 7)
BB8 = ArrDeviceIds_ExcelKR3(i, 8)
BB9 = ArrDeviceIds_ExcelKR3(i, 9)
BB10 = ArrDeviceIds_ExcelKR3(i, 10)

'-------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AI 08 052-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AI08_052-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AI08_052-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-1000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-1000AI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-1000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-1000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-1000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-1000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If
'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AI 16 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AI16_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AI16_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_AI16_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-1001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-1001AI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-1001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-1001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-1001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-1001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
	
	name = "-1001W2"
	deviceDesignation = BB10
	Call Pereimenovanie
	
End If


'-------------------------------------------------------------------------------------------------
If BB3 = "R500 AO 08 011-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
	
	searchSheetName = SheetName
	sheetId = sheet.Search( moduleId, searchSheetName )
	
	Call Include ("Контроллер_R500_AO08_052-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
	Call Include ("Контроллер_R500_AO08_052-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	name = "-2000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-2000AO1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-2000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-2000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-2000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-2000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If

'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DI 32 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DI32_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_012-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-3001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-3001DI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-3001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-3001KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-3001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-3001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-3001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If


'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DI 32 013-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DI32_013-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DI32_013-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-3001CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-3001DI1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-3001X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-3001KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-3001XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-3001XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-3001W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If

'-------------------------------------------------------------------------------------------------
If BB3 = "R500 DO 32 012-000-AAA" Then
	DocumentTypeAttr = ".DOCUMENT_TYPE" 	'атрибут - тип документа
	NameDocument = "Naimenovanie_Documenta" 'атрибут - наименование документа
	SheetFormat = "Format" 					'атрибут - формат
	NumSheets = "_numSheetsInDoc" 			'атрибут - кол-во листов
	
	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_0.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента
			Call Include ("Контроллер_R500_DO32_012-000-AAA_1.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_2.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_3.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента

	Total_sheet = job.GetSheetCount()
	SheetName = Total_sheet + 1
	Sheet.Create 0, sheetName, "Формат_А3_гор_2лист", 0, 1
	Sheet.SetAttributeValue DocumentTypeAttr, "Схема электрическая принципиальная"
		searchSheetName = SheetName
		sheetId = sheet.Search( moduleId, searchSheetName )
			Call Include ("Контроллер_R500_DO32_012-000-AAA_4.vbs") ' функция выполняет указанный файл (скрипт) - размещение фрагмента


	name = "-4000CH1"
	deviceDesignation = BB0
	Call Pereimenovanie_Chassi
	
	name = "-4000DO1"
	deviceDesignation = BB2
	Call Pereimenovanie
	
	name = "-4000X1"
	deviceDesignation = BB4
	Call Pereimenovanie
	
	name = "-4000KL"
	deviceDesignation = BB6
	Call Pereimenovanie_KL
	
	name = "-4000XT1"
	deviceDesignation = BB7
	Call Pereimenovanie
	
	name = "-4000XT2"
	deviceDesignation = BB8
	Call Pereimenovanie
	
	name = "-4000W1"
	deviceDesignation = BB9
	Call Pereimenovanie
End If
Next 
End If
' ======================================================================================================
































'===============================================================================================


' ===============================================================
' Заполнение штампа на листах
' ===============================================================
'Call Include("Верса_200_Proekt.vbs")


' ===============================================================
' ВСТАВКА СТРАНИЦ ДЛЯ Э3
' ===============================================================
'Call Vstavka_stranits



'

app.PutInfo 0, "==============================================================="
app.PutInfo 0, "ГЕНЕРАЦИЯ ЗАВЕРШЕНА"
app.PutInfo 0, "==============================================================="





' ===============================================================
' Процедура подключения кода
' ===============================================================
Sub Include (ByVal fileName)
	' Переменные для работы с файлами
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	' Запрос текущй папки скрипта
	Dim thisFolder: thisFolder = fso.GetParentFolderName(WScript.ScriptFullName)
	' Полный путь до нужного файлами
	Dim fileFullName
	fileFullName = fso.BuildPath(thisFolder, fileName)
	' Проверка файла
	If (fso.FileExists(fileFullName)) Then
		' Выполняем его
'		Call ExecuteGlobal(fso.OpenTextFile(fileFullName).ReadAll())
		Call ExecuteGlobal(fso.OpenTextFile(fileFullName, 1, False, -2 ).ReadAll())
	Else
		' Вывод сообщения об ошибке
		Call MsgBox("Ошибка открытия файла " & fileFullName & ". Файла не существует!", 16, "Ошибка открытия файла")
		' Очистка объекта
		Set fso = Nothing
		' Выход из выполнения скрипта
		WScript.Quit
	End If
	' Очистка объекта
	Set fso = Nothing
End Sub




Set sym = Nothing
Set component = Nothing
Set symbol = Nothing
Set slot = Nothing
Set devicePin = Nothing
Set connection = Nothing
Set pin = Nothing
Set device = Nothing
Set dev = Nothing
Set sheet = Nothing
Set job = Nothing
Set e3 = Nothing
wscript.Quit