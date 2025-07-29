' =================
' Настройки скрипта
' =================

' Фильтрация
' Работать только по активному листу (приоритет 1)
Const flagModeDoOnlyActiveSheet = False
' Выполнять автоматизацию для выделенных листов (приоритет 2)
Const flagModeDoSelectedSheets = True
' Выполнять автоматизацию для всех объектов в выделенных папках (приоритет 2)
Const flagModeDoSelectedFolders = True

' Работать в обработке только в одним листом, то есть если количество листов превысит один, то скрипт выдаст ошибку, если опция включена
Const flagModeDoSelectedOnlyOneSheet = False

'--------------------------------------------------------------

' Режим отладки исходных данных ("да" = true, "нет" - false)
Const deBugInputMode = false
' Режим отладки процесса обработки данных ("да" = true, "нет" - false)
Const deBugProcessMode = false
' Режим отладки вывода данных ("да" = true, "нет" - false)
Const deBugOutputMode = false
' Режим вывода сообщений в окно результатов для работы автоматизации для конечного пользователя
Const showMessageMode = True

' если свой вариант, пользовательский (настройки внутри тела программы), то значение = 1
' если настройки выполняются в соответствии с цветом провода/жилы, то значение = 2
' если настройки выполняются в соответствии с именем цепи провода/жилы, то значение = 0 
Const modeManualSetting = 2
' устанавливаем флаг, если настройки соответствия выполнять по значению имени цепи ТОЛЬКО ДЛЯ ЦЕПЕЙ БЕЗ НАЗНАЧЕННЫХ ПРОВОДНИКОВ, то true
Const flagSetColourFromSignalName = True
' устанавливаем флаг, если настройки соответствия выполнять по параметрам проводов этой цепи -  ТОЛЬКО ДЛЯ ЦЕПЕЙ БЕЗ НАЗНАЧЕННЫХ ПРОВОДНИКОВ, то true
Const flagSetColourFromCoreForSignalName = False

' толщина линии связи по-умолчанию
Const netSegmentLineWidthForOneSignal = 0.5
' добавление утолщения при прохождении двух или более цепей в сегменте
Const netSegmentLineWidthForTwoSignal = 0.8

' ===============================================================================================
' Процедура с настройкой по атрибутам кабеля или проводника (при modeManualSetting = 1)
Sub SubGetColourByUserSetting(ByVal cor, ByVal cab, ByVal job, ByRef netSegmentLineColour, ByRef netSegmentLineStyle, ByRef netSegmentLineWidth)
	' Определяем условия по которым будет назначаться цвет

	' Установка толщины в соответствии с настройкой	по умолчанию
	netSegmentLineWidth = netSegmentLineWidthForOneSignal
	' Вариант 1
	If cab.IsWiregroup Then
		' Условие 1 - если провод то линия связи должна иметь цвет 15
		netSegmentLineColour = "15"
		netSegmentLineStyle = "1"
		If netCoreColourDescription = "желто-зеленый" Then
			netSegmentLineColour = "58"
			netSegmentLineStyle = "5"
		End If
		' -----------------------	
	ElseIf cab.GetAttributeValue("Тех. описание 1") = "1" Then
		' Условие 2 - если провод то линия связи должна иметь цвет 17
		netSegmentLineColour = "17"
		netSegmentLineStyle = "1"
		netSegmentLineWidth = netSegmentLineWidthForTwoSignal
		' -----------------------
	ElseIf cab.GetAttributeValue("Тех. описание 1") = "2" Then
		' Условие 3 - если провод то линия связи должна иметь цвет 18
		netSegmentLineColour = "18"
		netSegmentLineStyle = "1"
		netSegmentLineWidth = netSegmentLineWidthForTwoSignal
		' -----------------------
	Else
		' Условия не выполняются - линия связи должна иметь цвет чёрный, т.е. 0
		netSegmentLineColour = "0"
		netSegmentLineStyle = "1"
		' -----------------------				
	End If
End Sub


' ===============================================================================================
' Процедура назначения цветов по цвету проводника (при modeManualSetting = 2)
Sub SubGetColourFromCore(ByVal coreColourDescription, ByRef netSegmentLineColour, ByRef netSegmentLineStyle, ByRef netSegmentLineWidth)
	' Установка значений по умолчанию
	netSegmentLineColour = "256"
	netSegmentLineStyle = "1"
	' Установка толщины в соответствии с настройкой	по умолчанию
	netSegmentLineWidth = netSegmentLineWidthForOneSignal
	' Условия соответствия
	If coreColourDescription = "черный" Then
		netSegmentLineColour = "0"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "белый" Then
		netSegmentLineColour = "19"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "красный" Then
		netSegmentLineColour = "13"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "зеленый" Or coreColourDescription = "зелёный" Then
		netSegmentLineColour = "14"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "коричневый" Then
		netSegmentLineColour = "30"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "голубой" Then
		netSegmentLineColour = "16"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "оранжевый" Then
		netSegmentLineColour = "57"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "желтый" Or coreColourDescription = "жёлтый" Then
		netSegmentLineColour = "15"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "фиолетовый" Then
		netSegmentLineColour = "5"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "серый" Then
		netSegmentLineColour = "12"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "розовый" Then
		netSegmentLineColour = "17"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "золотистый" Then
		netSegmentLineColour = "63"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "бирюзовый" Then
		netSegmentLineColour = "200"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "серебристый" Then
		netSegmentLineColour = "8"
		netSegmentLineStyle = "1"
	ElseIf coreColourDescription = "желто-зеленый" Or coreColourDescription = "жёлто-зелёный" Or coreColourDescription = "зелено-желтый" Or coreColourDescription = "зелёно-жёлтый" Then
		netSegmentLineColour = "58"
		netSegmentLineStyle = "5"
	ElseIf coreColourDescription = "синий" Then
		netSegmentLineColour = "16"
		netSegmentLineStyle = "1"
	End If
End Sub

' ===============================================================================================
' Процедура назначения цветов по значению имени цепи (при modeManualSetting = 0 и flagSetColourFromSignalName = True)
Sub SubGetColourFromSignalName(ByVal signalName, ByRef netSegmentLineColour, ByRef netSegmentLineStyle, ByRef netSegmentLineWidth)
	' Установка значений по умолчанию
	netSegmentLineColour = "205"
	netSegmentLineStyle = "1"
	' Установка толщины в соответствии с настройкой	по умолчанию
	netSegmentLineWidth = netSegmentLineWidthForOneSignal
	' Условия соответствия
	If signalName = "2L1" Or signalName = "L1" Or signalName = "1L1" Then
		netSegmentLineColour = "0"
		netSegmentLineStyle = "1"
		netSegmentLineWidth = 2
	ElseIf signalName = "2L2" Or signalName = "L2" Or signalName = "1L2" Then
		netSegmentLineColour = "38"
		netSegmentLineStyle = "1"
		netSegmentLineWidth = 2
	ElseIf signalName = "2L3" Or signalName = "L3" Or signalName = "1L3" Then
		netSegmentLineColour = "11"
		netSegmentLineStyle = "1"
		netSegmentLineWidth = 2
	ElseIf signalName = "N" Or signalName = "1N" Or signalName = "2N" Or signalName = "3N" Or signalName = "N1" Or signalName = "N2" Or signalName = "N.A" Or signalName = "N2_1" Or signalName = "N2_2" Then
		netSegmentLineColour = "160"
		netSegmentLineStyle = "4"
		netSegmentLineWidth = 1
	ElseIf signalName = "PE" Or signalName = "PEN" Then
		netSegmentLineColour = "58"
		netSegmentLineStyle = "5"
		netSegmentLineWidth = 1
	ElseIf signalName = "L1.A" Or signalName = "-24D" Then
		netSegmentLineColour = "0"
		netSegmentLineStyle = "4"
		netSegmentLineWidth = 1	
	ElseIf signalName = "+24" Or signalName = "+24_1" Or signalName = "+24_2" Or signalName = "+24_3" Or signalName = "+24_4" Or signalName = "A" Then
		netSegmentLineColour = "10"
		netSegmentLineStyle = "4"
		netSegmentLineWidth = 1	
	ElseIf signalName = "-24" Or signalName = "B" Or signalName = "-24_1" Or signalName = "-24_2" Or signalName = "-24_3" Or signalName = "-24_4" Then
		netSegmentLineColour = "12"
		netSegmentLineStyle = "4"
		netSegmentLineWidth = 1
    ElseIf signalName = "L24" Or signalName = "N24" Or signalName = "N24_1" Or signalName = "24AC" Or signalName = "+24D" Then
		netSegmentLineColour = "25"
		netSegmentLineStyle = "4"
		netSegmentLineWidth = 1		
	End If
End Sub
' ===============================================================================================
' =============================
' Объявление глобальных переменных
Dim app, appId, appVersion, jobId, jobName, resultMsg

' Запускаем скрипт
Call SubStartScript()

' Переходим к основному телу скрипта
Call SubMainProcessScript()

' Процедура выхода
Call SubFinishScript(False)

' Работа скрипта завершена
' =====================================
' ===============================================================================================
' ===============================================================================================

' ==========
' Процедуры:
' ==========
' =========================
' Процедура запуска скрипта
' =========================
Sub SubStartScript()
	' Создаём объект app
	Set app = CreateObject ("CT.Application")
	' Идентификатор процесса
	appId = app.GetId()
	' Если нет открытого приложения E3.series то
	If (appId = 0) Then
		' Выводим сообщение об ошибке
		msgbox("Нет запущенного приложения E3.series. Выход!")
		' Выходим
		Call SubFinishScript(True)
	End If
	' Текущая версия E3.series
	appVersion = CInt(app.GetVersion())
	' Создаём объект job
	Dim job
	Set job = app.CreateJobObject ()
	' Определяем идентификатор проекта
	jobId = job.GetId()
	' Проверка открытого проекта
	If (jobId = 0) Then
		' Выводим сообщение об ошибке
		resultMsg = app.PutError(1, "Нет открытого проекта!")
		' Выходим
		Call SubFinishScript(True)
	End If

	' Считываем имя проекта
	jobName = job.GetName
	If showMessageMode Then resultMsg = resultMsg = app.PutMessage("")
	If showMessageMode Then resultMsg = app.PutMessage("")
	If showMessageMode Then resultMsg = app.PutMessage("")
	If showMessageMode Then resultMsg = app.PutInfo(0, "=========================================")
	If showMessageMode Then resultMsg = app.PutInfo(0, "Начало работы автоматического назначения цветов линиям связи '" & jobName & "'", job.GetId)

	' Очистка
	Set job = Nothing

End Sub

' ===================================
' Процедура окончания
Sub SubFinishScript(ByVal flagExitScript)
	' Создаём объект job
	Dim job
	Set job = app.CreateJobObject ()

	' Проверка флага завершения работы
	If flagExitScript Then
		' Выход с ошибкой
		If showMessageMode Then resultMsg = app.PutInfo(0, "=========================================")
		If showMessageMode Then resultMsg = app.PutError(0, "Работа автоматизации не выполнена! Выход...")
	Else
		' Успешный выход
		If showMessageMode Then resultMsg = app.PutInfo(0, "=========================================")
		If showMessageMode Then resultMsg = app.PutInfo(0, "Работа автоматизации успешно завершена! Выход...")
	End If
	' Отчищаем объекты
	Set job = Nothing
	Set app = Nothing

	'Работа завершена, Выходим
	wscript.quit
End Sub

' ===============================
' Процедура рабочего тела скрипта

Sub SubMainProcessScript()
	' Создаём объект job
	Dim job
	Set job = app.CreateJobObject ()

	' Проверка правильности настроек
	' Какие настройки включены

	If flagModeDoOnlyActiveSheet Then
		If showMessageMode Then resultMsg = app.PutInfo(0, "Включена опция работы только с активным листом в обработке!")
	Else
		If showMessageMode Then resultMsg = app.PutInfo(0, "Отключена опция работы только с активным листом в обработке!")

		If flagModeDoSelectedSheets Then
			If showMessageMode Then resultMsg = app.PutInfo(0, "Включена опция работы с выделенными листами!")
		Else
			If showMessageMode Then resultMsg = app.PutInfo(0, "Отключена опция работы с выделенными листами!")
		End If
		If flagModeDoSelectedFolders Then
			If showMessageMode Then resultMsg = app.PutInfo(0, "Включена опция работы с выделенными папками!")
		Else
			If showMessageMode Then resultMsg = app.PutInfo(0, "Отключена опция работы с выделенными папками!")
		End If
	End If

	If flagModeDoSelectedOnlyOneSheet Then
		If showMessageMode Then resultMsg = app.PutInfo(0, "Включена опция работы только с одним выделенным листом в обработке!")
	Else
		If showMessageMode Then resultMsg = app.PutInfo(0, "Отключена опция работы только с одним выделенным листом в обработке!")
	End If


	' ------------------------------------------------------
	If showMessageMode Then resultMsg = app.PutInfo(0, "----------------------------------------")

	' ===============================================
	' Создаём объект tree
	Dim tree, treeId
	Set tree = job.CreateTreeObject
	' Устанавливаем активное дерево текущим
	treeId = tree.SetId(job.GetActiveTreeId())
	' Активный лист (идентификатор)
	activeSheetId = job.GetActiveSheetId

	' ===============================================
	' определяем количественные переменные массивов
	Dim shtId, activeSheetId
	' определяем словарь для хранения данных
	Dim dictSheetIds, dictSheetCnt
	Set dictSheetIds = CreateObject("Scripting.Dictionary")
	dictSheetCnt = 0
	' Словарь для жил
	Dim dictCoreIds, dictCoreCnt
	Set dictCoreIds = CreateObject("Scripting.Dictionary")
	dictCoreCnt = 0

	' ===========================
	' Сбор исходных данных
	'-------------------------------------------
	' Запрос данных в зависимости от включенных различных опций 
	If (flagModeDoOnlyActiveSheet) Then
		' Если так, то работаем только по активному листу то
		' Переходим к процедуре
		dictSheetCnt = FunGetDictSheetIds(activeSheetId, dictSheetIds, dictCoreIds)
		If deBugInputMode Then resultMsg = app.PutInfo(0, "Автоматизация выполняется по активному листу...", activeSheetId)
	Else
		' иначе поиск в других местах
		' Поиск выделенных в дереве листов
		cntTreeSelectedSheet = tree.GetSelectedSheetIds(treeSelectedSheetIds)
		If deBugInputMode Then resultMsg = app.PutInfo(0, "Количество выделенных листов в дереве = " & cntTreeSelectedSheet)
		' Проверка наличия выделенных в дереве символов
		If (cntTreeSelectedSheet > 0) And flagModeDoSelectedSheets Then
			' Перебор символов 
			For iTreeSelectedSheet = 1 To cntTreeSelectedSheet
				shtId = treeSelectedSheetIds(iTreeSelectedSheet)
				dictSheetCnt = FunGetDictSheetIds(shtId, dictSheetIds, dictCoreIds)
			Next
		End If

		' Поиск листов в выделенных папках
		cntTreeSelectedSheetByFolder = tree.GetSelectedSheetIdsByFolder(treeSelectedSheetIdsByFolder)
		If deBugInputMode Then resultMsg = app.PutInfo(0, "Количество выделенных листов в папке дерева = " & cntTreeSelectedSheetByFolder)
		' Проверка наличия выделенных в дереве изделий 
		If (cntTreeSelectedSheetByFolder > 0) And flagModeDoSelectedFolders Then
			' Перебор выбранных изделий 
			For iTreeSelectedSheetByFolder = 1 To cntTreeSelectedSheetByFolder
				shtId = treeSelectedSheetIdsByFolder(iTreeSelectedSheetByFolder)
				dictSheetCnt = FunGetDictSheetIds(shtId, dictSheetIds, dictCoreIds)
			Next
		End If
	End If

	' Если не нашли выделенные листы, то берём активный
	If ((activeSheetId > 0) And (dictSheetCnt = 0)) Then
		' Переходим к процедуре
		dictSheetCnt = FunGetDictSheetIds(activeSheetId, dictSheetIds, dictCoreIds)
		If deBugInputMode Then resultMsg = app.PutInfo(0, "Автоматизация выполняется по активному листу...", activeSheetId)
	End If

	If showMessageMode Then resultMsg = app.PutInfo(0, "----------------------------------------")
	If showMessageMode Then resultMsg = app.PutInfo(0, "Количество листов в обработке = " & dictSheetCnt)

	' ================================
	' Работа со всеми найденными листами
	If ((dictSheetCnt > 0) And (Not flagModeDoSelectedOnlyOneSheet)) Or ((dictSheetCnt = 1) And (flagModeDoSelectedOnlyOneSheet)) Then
		If showMessageMode Then resultMsg = app.PutMessage(vbTab & "Выполняется сбор данных о листах и их сортировка...")
		' Создание массива для хранения всех данных листов
		ReDim arrSortSheetIds(dictSheetCnt - 1, 4)
		' Заполнение массива
		For shti = 0 To dictSheetCnt - 1
			' Установка идентификатора
			shtId = dictSheetIds.Keys()(shti)
			' Заполнение массива
			arrSortSheetIds(shti, 0) = shtId
			arrSortSheetIds(shti, 1) = dictSheetIds.Item(shtId).Assignment
			arrSortSheetIds(shti, 2) = dictSheetIds.Item(shtId).Location
			arrSortSheetIds(shti, 3) = dictSheetIds.Item(shtId).DOCUMENTTYPE
			arrSortSheetIds(shti, 4) = dictSheetIds.Item(shtId).Name

		Next

		' Сортировка листов
		Call subSortArrayByIndexEx(arrSortSheetIds, array(2, 3, 4, 5), array(2, 2, 2, 2), array(0, 0, 0, 0))

		' ------------------------------------------
		' Перебор массива листов
		For shti = 0 To dictSheetCnt - 1
			' Установка идентификатора
			shtId = dictSheetIds.Keys()(shti)
			If showMessageMode Then resultMsg = app.PutInfo(0, "----------------------------------------")
			If showMessageMode Then resultMsg = app.PutInfo(0, "Работаем с листом = " & dictSheetIds.Item(shtId).Assignment & " " & dictSheetIds.Item(shtId).Location & " / " & dictSheetIds.Item(shtId).DOCUMENTTYPE & " / лист " & dictSheetIds.Item(shtId).Name & " - формат листа: " & dictSheetIds.Item(shtId).Format, shtId)

			' Проверка количества сегментов сетей на листе
			If (dictSheetIds.Item(shtId).DictNetSegmentIds.Count > 0) Then
				' Вывод сообщения
				If showMessageMode Then resultMsg = app.PutInfo(0, vbTab & "Выполняется сбор данных сегментов графических сетей...")
				' Перебор
				For Each nsegId In dictSheetIds.Item(shtId).DictNetSegmentIds.Keys()
					' Переход к работе с сетями на листе
					Call subWorkShtNetSegmentIds(nsegId, dictCoreIds)
				Next
			Else
				' Если не найдены исходные данные, то выводим сообщение и выходим с ошибкой...
				If showMessageMode Then resultMsg = app.PutWarning(0, vbTab & "На текущем листе в проекте нет графических сетей!")
			End If
		Next

		' Очистка объектов
		Erase arrSortSheetIds

	ElseIf ((dictSheetCnt = 0) And ((flagModeDoOnlyActiveSheet) Or (flagModeDoSelectedSymbols) Or (flagModeDoTriggerAfterModifySymbol))) Then
		If showMessageMode Then resultMsg = app.PutError(1, "В проекте '" & jobName & "' не найден активный лист!", job.GetId)
		' Выходим
		Call SubFinishScript(True)
	Else
		If ((Not flagModeDoSelectedSymbols) And (Not flagModeDoTriggerAfterModifySymbol)) Then
			If showMessageMode Then resultMsg = app.PutError(1, "Нет выделенных в дереве листов в проекте '" & jobName & "'", job.GetId)
			' Выходим
			Call SubFinishScript(Not flagExitScript)
		End If
	End If

	If ((dictSheetCnt > 1) And (flagModeDoSelectedOnlyOneSheet)) Then
		If showMessageMode Then resultMsg = app.PutError(1, "В обработке должен быть только один лист!")
		' Выходим
		Call SubFinishScript(True)
	End If

	' Очистка
	Set tree = Nothing
	Set treeId = Nothing
	
	Set dictSheetIds = Nothing
	Set dictSheetCnt = Nothing
	Set dictCoreIds = Nothing
	Set dictCoreCnt = Nothing

	' Очистка
	Set job = Nothing

End Sub

' ====================================================================
' ====================================================================
' Процедура работы с телом скрипта
Sub subWorkShtNetSegmentIds(ByVal itemId, ByRef dictCoreIds)
	' Проверка
	If (itemId > 0) Then
		' Создаём объект job
		Dim job
		Set job = app.CreateJobObject ()

		' Создание объектов
		Dim nseg, nsegId, nsegLineColour, nsegLineStyle, nsegLineWidth, nsegSignalName
		Dim corId, corSignalName
		Set nseg = job.CreateNetSegmentObject()
		' Устанавливаем текущий идентификатор
		nsegId = nseg.SetId(itemId)
		' Проверка
		If (nsegId > 0) Then
			' Цвет будет авто = 256 по умолчанию
			nsegLineColour = "256"
			nsegLineStyle = "1"
			' Определяем значение толщины по умолчанию
			nsegLineWidth = netSegmentLineWidthForOneSignal
			' запрос имени цепи в сегменте
			nsegSignalName = nseg.GetSignalName()
			' Включать назначение при включении соответствующего флага (работает только для не назначенных проводниками цепей)
			If (flagSetColourFromSignalName) Then
				' запрос соответствия цвета в зависимости от имени цепи
				Call SubGetColourFromSignalName(nsegSignalName, nsegLineColour, nsegLineStyle, nsegLineWidth)
			End If

			' ------------------------------------
			' Включать назначение при включении соответствующего флага (работает только для не назначенных проводниками цепей) определённого с учётом цвета остальных жил
			If (flagSetColourFromCoreForSignalName) Then
				' Перебор всех жил
				For Each corId In dictCoreIds.Keys()
					' Имя цепи жилы
					corSignalName = dictCoreIds.Item(corId).SignalName
					' Поиск жилы с аналогичной цепью
					If (nsegSignalName = corSignalName) Then
						' Жила с такой цепью найдена - назначаем параметры по этой жиле
						nsegLineColour = dictCoreIds.Item(corId).NetSegmentColour
						nsegLineStyle = dictCoreIds.Item(corId).NetSegmentLineStyle
						nsegLineWidth = dictCoreIds.Item(corId).NetSegmentLineWidth
						' Выход из цикла
						Exit For
					End If
				Next
			End If

			' ------------------------------------
			' Запрос жил в каждом сегменте
			nsegCoreCnt = nseg.GetCoreIds(nsegCoreIds)
			' Проверка количества жил в сегменте
			If (nsegCoreCnt > 0) Then
				' Создание флагов
				Dim flagDifferentByColour, flagDifferentByStyle, flagDifferentBySignal
				' Установка флагов
				flagDifferentByColour = False
				flagDifferentByStyle = False
				flagDifferentBySignal = False

				' перебираем жилы
				For cori = 1 To nsegCoreCnt
					' Текущая жила 
					corId = nsegCoreIds(cori)
					' Поиск жилы в словаре
					If (dictCoreIds.Exists(corId)) Then
						' Жила найдена - назначаем параметры по этой жиле
						nsegLineColour = dictCoreIds.Item(corId).NetSegmentColour
						nsegLineStyle = dictCoreIds.Item(corId).NetSegmentLineStyle
						nsegLineWidth = dictCoreIds.Item(corId).NetSegmentLineWidth
						corSignalName = dictCoreIds.Item(corId).SignalName
						' --------------------
						' Если в сегменте много проводов разного соответствующего цвета
						' Фиксируем значение цвета, определённого для первой жилы в сегмента
						If cori = 1 Then
							' значение цвета для первого сегмента
							nsegColourFirst = nsegLineColour
							' Заполняем значение типа линии
							nsegLineStyleFirst = nsegLineStyle
							' Заполняем значение имени цепи жилы
							corSignalNameFirst = corSignalName
						End If

						' Выполняем сравнение цвета для каждой жилы с цветом первой жилы и её типом
						If (nsegColourFirst <> nsegLineColour) Or (nsegLineStyleFirst <> nsegLineStyle) Then
							' определяем, что в сегменте ошибка и красим его в авто, т.е. в 256
							nsegLineColour = "256"
							nsegLineStyle = "1"
							' Установка флагов для выхода из цикла
							flagDifferentByColour = True
							flagDifferentByStyle = True
						End If
						' Проверка имени цепи
						If (corSignalNameFirst <> corSignalName) Then
							' Задаём утолщение
							nsegLineWidth = netSegmentLineWidthForTwoSignal
							' Установка флагов для выхода из цикла
							flagDifferentBySignal = True
						End If
						' Проверка флагов
						If (flagDifferentByColour And flagDifferentByStyle And flagDifferentBySignal) Then
							' Выход из цикла
							Exit For
						End If
					End If
				Next
				' Очистка
				Set flagDifferentByColour = Nothing
				Set flagDifferentByStyle = Nothing
				Set flagDifferentBySignal = Nothing
			End If

			' ------------------------------------
			' Назначение цветов и типов линий линиям связи
			' ------------------------------------
			' Стиль назначается до цвета 
			nsegSetLineStyleResult = nseg.SetLineStyle(nsegLineStyle)
			' Возможность назначения цвета, зависит от типа линии, поэтому цвет линии назначается после
			segSetLineColourResult = nseg.SetLineColour(nsegLineColour)

			If deBugOutputMode Then resultMsg = app.PutError(0, vbTab & "netSegmentId = " & nsegId & ",  netSegmentLineColour = " & nsegLineColour & ", netSegmentLineStyle = " & nsegLineStyle, nsegId)

			' -------------------------------------------------------------------------------
			' Добавляем обработку утолщения линии связи при прохождении нескольких цепей (двух и более) в линии связи
			nsegSetLineWidthResult = nseg.SetLineWidth(nsegLineWidth)
			If deBugOutputMode Then resultMsg = app.PutError(0, vbTab & "netSegmentLineWidth = " & nsegLineWidth, nsegId)

			'Проверка версии E3.series
			If (appVersion >= 2018) Then
				' Перевод в RGB
				jobRGBValueRet = job.GetRGBValue(nsegLineColour, rColour, gColour, bColour)
				' Вывод сообщений
				If showMessageMode Then resultMsg = app.PutInfoEx(0, vbTab & "Успешно выполнено назначение типа, цвета и толщины для линии связи -" & nsegId & "-...", nsegId, rColour, gColour, bColour)
			Else
				' Вывод сообщений
				If showMessageMode Then resultMsg = app.PutInfo(0, vbTab & "Успешно выполнено назначение типа, цвета и толщины для линии связи -" & nsegId & "-...", nsegId)
			End If


		End If

		' Очистка
		Set nseg = Nothing
		Set nsegId = Nothing
		Set nsegLineColour = Nothing
		Set nsegLineStyle = Nothing
		Set nsegLineWidth = Nothing
		Set nsegSignalName = Nothing
		Set corId = Nothing
		Set corSignalName = Nothing
		' Очистка
		Set job = Nothing
	End If

End Sub
' =====================================

' ==============================================================
' Вспомогательная процедура для выполнения стандартных сортировок
Sub subSortArrayByIndexEx(ByRef arraySort, ByVal columnSort, ByVal parameterSort, ByVal directionSort)
	' Поверка, что на входе у нас действительно массив
	If (IsArray(arraySort)) Then
		' Проверка
		If (IsArray(columnSort) Or IsArray(parameterSort) Or IsArray(directionSort)) Then
			' Проверка
			If (Not IsArray(columnSort) Or Not IsArray(parameterSort) Or Not IsArray(directionSort)) Then
				' Вывод сообщения
				resultMsg = app.PutError(0, "Не верная конфигурация массива опций сортировки!")
				' Выходим
				Call SubFinishScript(True)
			End If

			' Запрос количества
			columnSortZise = UBound(columnSort)
			parameterSortZise = UBound(parameterSort)
			directionSortZise = UBound(directionSort)
			' Проверка
			If (columnSortZise <> parameterSortZise) Or (columnSortZise <> directionSortZise) Then
				' Вывод сообщения
				resultMsg = app.PutError(0, "Количество параметров в массиве опций сортировки не одинаково!")
				' Выходим
				Call SubFinishScript(True)
			End If
			' Создаём управляющий массив
			'---------------------------
			ReDim optionsSort(columnSortZise, 2)
			' Перебор индексов
			For opti = 0 To columnSortZise
				' Заполнение
				optionsSort(opti, 0) = Trim(columnSort(opti))
				optionsSort(opti, 1) = Trim(parameterSort(opti))
				optionsSort(opti, 2) = Trim(directionSort(opti))
			Next
		Else
			' Создаём управляющий массив
			'---------------------------
			ReDim optionsSort(0, 2)
			' 1 = Сортировка по 1-й колонке (начинается с единицы всегда)
			optionsSort(0, 0) = Trim(columnSort)
			' Параметр отвечающий за вид сортировки - в данном случае используется инженерная = 2
			optionsSort(0, 1) = Trim(parameterSort)
			' Параметр отвечающий за возрастание (0) или убывание (1) 
			optionsSort(0, 2) = Trim(directionSort)
			'---------------------------
		End If

		'Выполнение сортировки
		Call app.SortArrayByIndexEx(arraySort, optionsSort)
		'Сортировка уровня жил закончена
		' После сортировки, нужно иметь ввиду, что данные могут сместиться из 0-го индекса в 1-й, если directionSort = 0,
		' из-за того, что ПОСЛЕДНИЙ ИНДЕКС ИСХОДНОГО МАССИВА ВСЕГДА ИМЕЕТ ПУСТОЕ ЗНАЧЕНИЕ, КОТОРОЕ ПЕРЕМЕТИТСЯ В НУЛЕВОЙ ИНДЕКС ИТОГОВОГО МАССИВА
	End If
End Sub
' ==============================================================

' ==============================================================
' Функция заполнения массива листов
Function FunGetDictSheetIds(ByVal itemId, ByRef dictSheetIds, ByRef dictCoreIds)
	' Если символ размещён то идентификатор листа не будет равен нулю
	If (itemId > 0) Then
		' Создаём объект job
		Dim job
		Set job = app.CreateJobObject ()
		' Создание объектов
		Dim sht, shtId
		Set sht = job.CreateSheetObject()
		' Установка текущего идентификатора
		shtId = sht.SetId(itemId)
		' Проверка типа вложенного листа
		If (sht.IsPanel) Or (sht.IsTopology) Or (sht.IsFormboard) Then
			shtId = sht.SetId(sht.GetParentSheetId())
		End If
		' Проверка
		If (shtId > 0) Then
			' Проверка в словаре
			If (Not dictSheetIds.Exists(shtId)) Then
				' Добавление в словарь
				Call dictSheetIds.Add(shtId, New classSheet)
				' Установка области видимости
				With dictSheetIds.Item(shtId)
					' Идентификатор
					.Id = shtId
					.Assignment = sht.GetAssignment()
					.Location = sht.GetLocation()
					.Name = sht.GetName()
					' Запрос типа документа
					.DOCUMENTTYPE = sht.GetAttributeValue(".DOCUMENT_TYPE")
					' Тип листа (символ в базе данных)
					.Format = sht.GetFormat()

					' Запрос net (цепей на текущем листе)
					shtNetCnt = sht.GetNetIds(shtNetIds)
					' Проверка
					If (shtNetCnt > 0) Then
						' Создание объектов
						Dim net, netId
						Set net = job.CreateNetObject()
						' Переход к работе с сетями
						For neti = 1 To shtNetCnt
							' Устанавливаем текущее значение
							netId = net.Setid(shtNetIds(neti))
							' ---------------------------------------
							' Запрос жил у графической цепи
							netCoreCnt = net.GetCoreIds(netCoreIds)
							' Проверка количества
							If (netCoreCnt > 0) Then
								' Выполняем перебор массива жил
								For cori = 1 To netCoreCnt
									' Переход к функции заполнения жил
									Call FunGetDictCoreIds(netCoreIds(cori), dictCoreIds)
								Next
							End If
							' ---------------------------------------
							' Запрос сегментов
							netNetSegmentCnt = net.GetNetSegmentIds(netNetSegmentIds)
							' Проверка количества
							If (netNetSegmentCnt > 0) Then
								' Создание объектов
								Dim nsegId
								' Перебор
								For nsegi = 1 To netNetSegmentCnt
									' Текущий сегмент сети
									nsegId = netNetSegmentIds(nsegi)
									' Проверка
									If (Not .DictNetSegmentIds.Exists(nsegId)) Then Call .DictNetSegmentIds.Add(nsegId, "")
								Next
								' Очистка
								Set nsegId = Nothing
							End If
						Next
					End If
				End With
				' Определяем в общем случае есть ли у листа региона (если лист является E3.Panel, E3.Formboard, топологией)
				shtEmbeddedSheetCnt = sht.GetEmbeddedSheetIds(shtEmbeddedSheetIds)
				' Проверка
				If (shtEmbeddedSheetCnt > 0) Then
					' выполняем переход от листа к региону
					For shti = 1 To shtEmbeddedSheetCnt
						Call FunGetDictSheetIds(shtEmbeddedSheetIds(shti), dictSheetIds, dictCoreIds)
					Next
				End If
			End If
		End If
		' Очистка
		Set sht = Nothing
		Set shtId = Nothing
		' Очистка
		Set job = Nothing
	End If
	' Возврат функции
	FunGetDictSheetIds = dictSheetIds.Count
End Function
' =============================================================
' Создание класса листов
Class classSheet
	' Создание объектов
	Public Id, Assignment, Location, Name, DOCUMENTTYPE, Format
	Public DictNetSegmentIds
	' --------------------------------------------------------------------
	' Конструктор класса
	Private Sub Class_Initialize()
		'MsgBox("class classSheet started")
		Set DictNetSegmentIds = CreateObject("Scripting.Dictionary")
	End Sub
	' --------------------------------------------------------------------
	' Диструктор класса
	Private Sub Class_Terminate()
		'MsgBox("class classSheet terminated")
		' Очистка объектов
		Set Id = Nothing
		Set Assignment = Nothing
		Set Location = Nothing
		Set Name = Nothing
		Set DOCUMENTTYPE = Nothing
		Set Format = Nothing
		DictNetSegmentIds.RemoveAll
		Set DictNetSegmentIds = Nothing
	End Sub
End Class
' ==============================================================
' Функция заполнения словаря жил
Function FunGetDictCoreIds(ByVal itemId, ByRef dictCoreIds)
	' Проверка
	If (itemId > 0) Then
		' Создаём объект job
		Dim job
		Set job = app.CreateJobObject ()
		' инициализируем переменные
		Dim cor, corId
		Set cor = job.CreatePinObject()
		Dim cab, cabId
		Set cab = job.CreateDeviceObject()
		' Установка идентификатора

		' инициализируем текущую жилу
		corId = cor.SetId(itemId)
		' Проверка
		' Проверка
		If (corId > 0) Then
			' Проверка
			If (Not dictCoreIds.Exists(corId)) Then
				' Добавление в словарь
				Call dictCoreIds.Add(corId, New classCore)
				' Установка
				With dictCoreIds.Item(corId)
					' Идентификатор
					.Id = corId
					' текущий кабель
					cabId = cab.SetId(corId)
					' Идентификатор кабеля
					.CableId = cabId
					' Запрос цвета провода
					.ColourDescription = cor.GetColourDescription()
					' Запрос цепи провода
					.SignalName = cor.GetSignalName()

					' ============================================================================
					' Создание объектов
					Dim nsegLineColour, nsegLineStyle, nsegLineWidth

					' определяем флаг настроек, если вручную определяются настройки, то флаг включен
					If modeManualSetting = 1 Then
						' -----------------------
						' Вариант 1 - по правилам пользователя
						Call SubGetColourByUserSetting(cor, cab, job, nsegLineColour, nsegLineStyle, nsegLineWidth)
					ElseIf modeManualSetting = 2 Then
						' -----------------------
						' Вариант 2 - по цвету жилы
						' запрос соответствия цвета от цвета проводника
						Call SubGetColourFromCore(.ColourDescription, nsegLineColour, nsegLineStyle, nsegLineWidth)

					Else
						' -----------------------
						' Вариант без определения настройки или 0, то - по имени цепи жилы/провода
						' запрос соответствия цвета в зависимости от имени цепи проводника
						Call SubGetColourFromSignalName(.SignalName, nsegLineColour, nsegLineStyle, nsegLineWidth)
					End If
					' ===========================================
					' Заполнение массива запись идентификатора и назначение параметров по цвету, типу и имени цепи жил внутри линии
					.NetSegmentColour = nsegLineColour
					.NetSegmentLineStyle = nsegLineStyle
					.NetSegmentLineWidth = nsegLineWidth
					' ----------------------------------
					' Очистка
					Set nsegLineColour = Nothing 
					Set nsegLineStyle = Nothing 
					Set nsegLineWidth = Nothing 

					' ============================================================================
				End With
			End If
		End If

		' Очистка
		Set cor = Nothing
		Set corId = Nothing
		Set cab = Nothing
		Set cabId = Nothing
		' Очистка
		Set job = Nothing
	End If
	' Возврат функции
	FunGetDictCoreIds = dictCoreIds.Count
End Function
' =============================================================
' Создание класса жил
Class classCore
	' Создание объектов
	Public Id, CableId, ColourDescription, SignalName, NetSegmentColour, NetSegmentLineStyle, NetSegmentLineWidth
	' --------------------------------------------------------------------
	' Конструктор класса
	Private Sub Class_Initialize()
		'MsgBox("class classCore started")
	End Sub
	' --------------------------------------------------------------------
	' Диструктор класса
	Private Sub Class_Terminate()
		'MsgBox("class classCore terminated")
		' Очистка объектов
		Set Id = Nothing
		Set CableId = Nothing
		Set ColourDescription = Nothing
		Set SignalName = Nothing
		Set NetSegmentColour = Nothing
		Set NetSegmentLineStyle = Nothing
		Set NetSegmentLineWidth = Nothing
	End Sub
End Class
' =============================================================
' Создание класса сегмента сети
Class classNetSegment
	' Создание объектов
	Public Id
	' --------------------------------------------------------------------
	' Конструктор класса
	Private Sub Class_Initialize()
		'MsgBox("class classNetSegment started")
	End Sub
	' --------------------------------------------------------------------
	' Диструктор класса
	Private Sub Class_Terminate()
		'MsgBox("class classNetSegment terminated")
		' Очистка объектов
		Set Id = Nothing
		
	End Sub
End Class