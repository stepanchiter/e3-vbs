Set app = CreateObject( "CT.Application" ) 
Set job = app.CreateJobObject()
Set device = job.CreateDeviceObject()
Set conductor = job.CreatePinObject()
Set pin = job.CreatePinObject()
Set dev = job.CreateDeviceObject()
Set Sig = Job.CreateSignalObject
set Cab = Job.CreateDeviceObject
set Cor = Job.CreatePinObject


call SheetGrid
call IndexSet
call Sort
call ReNameWire
call Sort


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
Const FORMAT = "#<.SHEET><.GRID>"

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


Sub ReNameWire ' имя провода по цвету с выбором режима

nCabs = Job.GetCableCount				' Количество кабелей в проекте
if nCabs = 0 then
	App.PutInfo 1, "No cables in project, exiting..."
	wscript.quit
end If

' Диалог выбора режима переименования
Dim userChoice
userChoice = InputBox("Выберите режим переименования:" & vbCrLf & vbCrLf & "1 - Фазная логика (проверка атрибута ""Класс соединения"")" & vbCrLf & "0 - Стандартная логика" & vbCrLf & vbCrLf & "Введите 1 или 0:", "Режим переименования проводов", "0")

' Проверяем корректность ввода
If userChoice = "" Then
	App.PutInfo 1, "Переименование отменено пользователем"
	Exit Sub
End If

If userChoice <> "1" And userChoice <> "0" Then
	App.PutInfo 1, "Некорректный выбор. Используется стандартная логика (0)"
	userChoice = "0"
End If

Dim usePhaseLogic
usePhaseLogic = (userChoice = "1")

App.PutInfo 0, "==========================================================="
If usePhaseLogic Then
	App.PutInfo 0, "Выбрана ФАЗНАЯ логика переименования"
Else
	App.PutInfo 0, "Выбрана СТАНДАРТНАЯ логика переименования"
End If
App.PutInfo 0, "==========================================================="

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
			 
			 ' Получаем атрибут "Класс соединения (Компоновка)"
			 ConnectionClass = Cor.GetAttributeValue("Класс соединения (Компоновка)")
			 
				' Переименовываем только провода, имя которых начинается с #
				If Left(WireName, 1) = "#" Then
					' Проверяем, есть ли уже новое имя для этого исходного имени
					If NameMapping.Exists(WireName) Then
						NewName = NameMapping(WireName)
					Else
						' Выбираем логику переименования в зависимости от пользовательского выбора
						If usePhaseLogic Then
							' Получаем атрибут "Класс соединения (Компоновка)" только если выбрана фазная логика
							ConnectionClass = Cor.GetAttributeValue("Класс соединения (Компоновка)")
							
							' Проверка фазности: если атрибут = "230/400V", используем фазную логику
							If ConnectionClass = "230/400V" Then
								' Фазное переименование
								If WireColor = "черный" Then
									NewName = "L1." & L1Counter
									L1Counter = L1Counter + 1
								ElseIf WireColor = "коричневый" Then
									NewName = "L2." & L2Counter
									L2Counter = L2Counter + 1
								ElseIf WireColor = "серый" Then
									NewName = "L3." & L3Counter
									L3Counter = L3Counter + 1
								' Синий провод остается как N
								ElseIf WireColor = "синий" Then
									NewName = "N" & NCounter
									NCounter = NCounter + 1	
								' Зелено-желтый провод остается как PE
								ElseIf WireColor = "зелено-желтый" Then
									NewName = "PE"								
								Else
									' Все остальные провода с общим счетчиком (просто номер)
									NewName = LCounter
									LCounter = LCounter + 1	
								End If
								
								App.PutInfo 0, "Фазное переименование: провод " & WireName & " (" & WireColor & ") с классом " & ConnectionClass
							Else
								' Если атрибут не 230/400V, используем общий счетчик для всех кроме N и PE
								If WireColor = "синий" Then
									NewName = "N" & NCounter
									NCounter = NCounter + 1	
								ElseIf WireColor = "зелено-желтый" Then
									NewName = "PE"								
								Else
									NewName = LCounter
									LCounter = LCounter + 1	
								End If
								
								App.PutInfo 0, "Провод без фазного класса: " & WireName & " (" & WireColor & ")"
							End If
						Else
							' Стандартная логика переименования (исходная)
							If WireColor = "черный" Then
								NewName = "L1." & L1Counter
								L1Counter = L1Counter + 1
							ElseIf WireColor = "коричневый" Then
								NewName = "L2." & L2Counter
								L2Counter = L2Counter + 1
							ElseIf WireColor = "серый" Then
								NewName = "L3." & L3Counter
								L3Counter = L3Counter + 1
							ElseIf WireColor = "синий" Then
								NewName = "N" & NCounter
								NCounter = NCounter + 1	
							ElseIf WireColor = "зелено-желтый" Then
								NewName = "PE"								
							Else
								NewName = LCounter
								LCounter = LCounter + 1	
							End If
							
							App.PutInfo 0, "Стандартное переименование: провод " & WireName & " (" & WireColor & ")"
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





Set dev = Nothing
Set pin = Nothing
Set conductor = Nothing
Set device = Nothing   
Set job = Nothing 
Set app = Nothing
Set Sig = Nothing
Set Cab = Nothing
Set Cor = Nothing