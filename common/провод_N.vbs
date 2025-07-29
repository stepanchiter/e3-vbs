' Задается: Номер цепи, для которой надо сделать соединения 
' Задается марка провода и обозначение
' 
' Марка провода и обозначение фиксированные 

Set app = CreateObject( "CT.Application" ) 
Set job = app.CreateJobObject()
Set symbol = job.CreateSymbolObject()
Set pin = job.CreatePinObject()
Set devicePin = job.CreatePinObject()
Set device = job.CreateDeviceObject()
Set connection = job.CreateConnectionObject()

Set dictPinIds1 = Nothing
Set dictPinIds1 = CreateObject("Scripting.Dictionary")

' Задаем провод и обозначение, задаем для каких изделий
wiregroupName = InputBox("Тип провода", "", "ПУГВнг(А)-LS")
If wiregroupName = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If
databaseWireName = InputBox("Сечение и цвет", "", "1х0.75(синий)")
If databaseWireName = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If
wireName = InputBox("Имя провода", "", "N")
If wireName = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If
signalName1 = InputBox("Имя цепи", "", "N")
If signalName1 = "" Then
    app.PutInfo 0, "Отменено пользователем."
    WScript.Quit
End If




app.PutInfo 0, "Создание и подключение провода"

namePozOboz = "-SF" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-QF" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-FU" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-KM" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-G" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-TV" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-A" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA_CPU
namePozOboz = "-QFD" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA_CPU
namePozOboz = "-V" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA_CPU
namePozOboz = "-KL" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-1KL" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-2KL" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-KT" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-KV" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-QS" ' Задаем для каких изделий

poiskpin = SOZDANIE_PINA
namePozOboz = "-SK" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-EL" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-M" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA

namePozOboz = "-HL" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-SB" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA
namePozOboz = "-SA" ' Задаем для каких изделий
poiskpin = SOZDANIE_PINA

namePozOboz = "-XT1"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT2"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT3"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT4"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT5"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT6"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT7"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT8"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT9"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT10"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT11"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT12"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT13"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT14"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT15"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT16"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT17"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT18"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT19"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT
namePozOboz = "-XT20"   ' Задаем для каких изделий
poiskpin = SOZDANIE_PINAXT

'----------------------------------
' Для контроллера
'----------------------------------
Function SOZDANIE_PINA_CPU

connectionCount = job.GetConnectionIds( connectionIds )        ' Находим контакты (pin's) для указанной цепи
deviceCount = job.GetDeviceIds( deviceIds )        ' Находим контакты (pin's) для указанной цепи

	For deviceIndex = 1 To deviceCount
		deviceId = device.SetId( deviceIds( deviceIndex ) )
		deviceName = device.GetName()
	
		If InStr(1, deviceName, namePozOboz, 1) Then
		pinCount = device.GetPinIds( pinIds )
		For pinIndex = 1 To pinCount
			pinId = pin.SetId( pinIds( pinIndex ) )
			pinName = pin.GetName()
			deviceId = device.SetId( pinId )
			deviceName = device.GetName()
			result2 = pin.GetSignalName()
				If result2 = signalName1 Then
					For connectionIndex = 1 To connectionCount
					connectionId = connection.SetId( connectionIds( connectionIndex ) )
					connectionName = connection.GetName()
					result3 = connection.GetSignalName()
						If result3 = signalName1 Then
							pinId = pin.SetId( pinIds( pinIndex ) )
							pinName = pin.GetName()
							dictPinIds1.Add pinId, deviceName
							app.PutInfo 0, "Имя цепи для контакта " & pinName & " ( " & pinId & " ) изделия " & deviceName & " ( " & deviceId & " ) - " & result2
						Exit For
						End If
					Next
				End If
		Next
		End If
	Next
		' Вывод сообщения
		app.PutInfo 0, deviceName & " - " & pinName & "( " & pinId & " )"
End Function



'----------------------------------
' Для изделия
'----------------------------------
Function SOZDANIE_PINA

connectionCount = job.GetConnectionIds( connectionIds )        ' Находим контакты (pin's) для указанной цепи
deviceCount = job.GetDeviceIds( deviceIds )        ' Находим контакты (pin's) для указанной цепи

Redim ArrDeviceIds(deviceCount-1, 2)
k = 0
	
	For deviceIndex = 1 To deviceCount
		deviceId = device.SetId( deviceIds( deviceIndex ) )
		deviceName = device.GetName()
	
		If InStr(1, deviceName, namePozOboz, 1) Then
		pinCount = device.GetPinIds( pinIds )
		For pinIndex = 1 To pinCount
			pinId = pin.SetId( pinIds( pinIndex ) )
			pinName = pin.GetName()
			deviceId = device.SetId( pinId )
			deviceName = device.GetName()
			result2 = pin.GetSignalName()
				If result2 = signalName1 Then
					For connectionIndex = 1 To connectionCount
					connectionId = connection.SetId( connectionIds( connectionIndex ) )
					connectionName = connection.GetName()
					result3 = connection.GetSignalName()
						If result3 = signalName1 Then
							pinId = pin.SetId( pinIds( pinIndex ) )
							pinName = pin.GetName()
							app.PutInfo 0, "Имя цепи для контакта " & pinName & " ( " & pinId & " ) изделия " & deviceName & " ( " & deviceId & " ) - " & result2

							'Заполнение массива для сортировки
								ArrDeviceIds(k, 0) = deviceName
								ArrDeviceIds(k, 1) = pinName
								ArrDeviceIds(k, 2) = pinId
								k = k + 1
						Exit For
						End If
					Next
				End If
		Next
		End If
	Next
	
Redim options (2, 2)
	' Первая колонка для сортировки
	options (0, 0) = 1 ' Номер колонки
	options (0, 1) = 2 ' Способ сортировки, 2 - инженерная
	' Вторая колонка для сортировки
	options (1, 0) = 3 ' Номер колонки
	options (1, 1) = 2 ' Способ сортировки, 2 - инженерная

	' Сортиовка массива
	app.SortArrayByIndexEx ArrDeviceIds, options
	app.PutMessage "После сортировки"
	
	' Перебор после сортировки
	For i=0 To k-1
		' Восстанавливаем переменные из массива
		deviceName = ArrDeviceIds(i, 0)
		pinName = ArrDeviceIds(i, 1)
		pinId = ArrDeviceIds(i, 2)
		dictPinIds1.Add pinId, deviceName

		' Вывод сообщения
		app.PutInfo 0, deviceName & " - " & pinName & "( " & pinId & " )"
	Next	
	
End Function


'----------------------------------
' Для клемм
'----------------------------------
Function SOZDANIE_PINAXT

connectionCount = job.GetConnectionIds( connectionIds )        ' Находим контакты (pin's) для указанной цепи
deviceCount = job.GetDeviceIds( deviceIds )        ' Находим контакты (pin's) для указанной цепи

deviceId = device.Search(namePozOboz, assignment, location)
deviceId = deviceId
deviceName = device.GetName()
sborka = device.IsAssembly()
klemmnik = device.IsTerminalBlock()

deviceName = namePozOboz
deviceCount = device.SearchAll( deviceName , deviceAssignment, deviceLocation, deviceIds )

' Создание массива сортировки
Redim ArrDeviceIds(deviceCount-1, 2)

	If deviceCount > 0 Then 
	k = 0
		For i = 1 To deviceCount
			deviceId = device.SetId( deviceIds( i ) )
			deviceName = device.GetName()
			result = device.GetPinIds( pinIds )
			
			If result = 0 Then
			Else

			For pinIndex = 1 To result
			
			pinId = pin.SetId( pinIds( pinIndex ) )
			pinName = pin.GetName()
			result2 = pin.GetSignalName()
				If signalName1 = result2 Then
					app.PutInfo 0, "Имя цепи для клеммы " & pinName & " ( " & pinId & " ) клеммника " & namePozOboz & " ( " & deviceId & " ) - " & result2
					'Заполнение массива для сортировки
						ArrDeviceIds(i-1, 0) = deviceName
						ArrDeviceIds(i-1, 1) = pinName
						ArrDeviceIds(i-1, 2) = pinId
						k = k + 1
						' Вывод сообщения
						app.PutInfo 0, deviceName & " - " & pinName & " - " & pinId
					Exit For
				Else
				End If
			Next
			End If
		Next
	End If

Redim options (2, 2)
	' Первая колонка для сортировки
	options (0, 0) = 2 ' Номер колонки
	options (0, 1) = 2 ' Способ сортировки, 2 - инженерная
	' Вторая колонка для сортировки
	options (1, 0) = 3 ' Номер колонки
	options (1, 1) = 2 ' Способ сортировки, 2 - инженерная

	' Сортиовка массива
	app.SortArrayByIndexEx ArrDeviceIds, options
	
	app.PutMessage "После сортировки"
	
	' Перебор после сортировки
	For i=0 To k-1
		' Восстанавливаем переменные из массива
		deviceName = ArrDeviceIds(i, 0)
		pinName = ArrDeviceIds(i, 1)
		pinId = ArrDeviceIds(i, 2)

		dictPinIds1.Add pinId, deviceId

		' Вывод сообщения
		app.PutInfo 0, deviceName & " - " & pinName & "( " & pinId & " )" & k
	Next
End Function

Arr = dictPinIds1.Items
For i=0 To dictPinIds1.Count-1
	app.PutInfo 0, "найдены:  " & dictPinIds1.Keys()(i)
	If i <> dictPinIds1.Count-1 Then
		firstTerminalPinId = dictPinIds1.Keys()(i)
		secondTerminalPinId = dictPinIds1.Keys()(i+1)
		iDPIN3 = SOZDANIE_PROVODA
	End If
Next


'==========================================
'Создание провода
'==========================================
Function SOZDANIE_PROVODA
If firstTerminalPinId>0 And secondTerminalPinId>0 Then
cableCount = job.GetCableIds( cableIds )
	For cableIndex = 1 To cableCount
		cableId = device.SetId( cableIds( cableIndex ) )
		cableName = device.GetName()
		isWireGroup = device.isWireGroup()
			If isWireGroup = 1 Then
				pinCount = device.GetPinIds( pinIds )
				result = pin.CreateWire( wireName, wiregroupName, databaseWireName, cableId, 0, 0 )
				If result = 0 Then
				app.PutError 0, "Ошибка создания провода " & wireName
				Else    
					wireName = pin.GetName()
					app.PutInfo 0, "Новый провод " & wireName & " , " & wiregroupName & " , " & databaseWireName & " , создан, ( " & cableId & " )"
				End If
			End If
	Next
	pin.SetEndPinId 1, firstTerminalPinId
	pin.SetEndPinId 2, secondTerminalPinId
	app.PutInfo 0, "Новый провод " & wireName & " , " & wiregroupName & " , " & databaseWireName & " , подключен, ( " & cableId & " )"
End If
End Function

Set dictPinIds1 = Nothing

app.PutInfo 0, " ==============================================================="

Set connection = Nothing
Set device = Nothing
Set devicePin = Nothing
Set pin = Nothing
Set symbol = Nothing
Set job = Nothing 
Set app = Nothing 