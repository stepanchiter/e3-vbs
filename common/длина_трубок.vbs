Set app = CreateObject( "CT.Application" ) 
Set job = app.CreateJobObject()
Set device = job.CreateDeviceObject()
Set conductor = job.CreatePinObject()
Set pin = job.CreatePinObject()
Set dev = job.CreateDeviceObject()
Set Sig = Job.CreateSignalObject()
set Cab = Job.CreateDeviceObject()
set Cor = Job.CreatePinObject()


'call SheetGrid
'call IndexSet
'call Sort
'call ReNameWire
'call Sort
call GetWireCrossSections


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




Sub GetWireCrossSections ' длина термоусаживаемых трубок для маркировки проводов
    nCabs = Job.GetCableCount ' Количество кабелей в проекте
    If nCabs = 0 Then
        App.PutInfo 1, "No cables in project, exiting..."
        wscript.quit
    End If

    ' Создаем массив для хранения сечений проводов
    Dim WireCrossSections()
    ReDim WireCrossSections(0) ' Инициализируем массив с нулевым размером

    cablecount = Job.GetCableIds(cableids)
    For i = 1 To cablecount
        Cab.SetId cableids(i)
        If Cab.IsWiregroup Then
            wircnt = Cab.GetPinIds(wirids)
            For j = 1 To wircnt
                Cor.SetId wirids(j)
                WireCrossSection = Cor.GetCrossSectionDescription ' Получаем сечение провода

                ' Добавляем сечение провода в массив
                If WireCrossSection <> "" Then
                    If UBound(WireCrossSections) = 0 And WireCrossSections(0) = "" Then
                        WireCrossSections(0) = WireCrossSection
                    Else
                        ReDim Preserve WireCrossSections(UBound(WireCrossSections) + 1)
                        WireCrossSections(UBound(WireCrossSections)) = WireCrossSection
                    End If
                End If
            Next
        End If
    Next

    ' Создаем массивы для разных сечений
    Dim A(), B(), C(), E()
    ReDim A(0)
    ReDim B(0)
    ReDim C(0)
    ReDim E(0)

    ' Разделяем сечения по заданным значениям
    For k = 0 To UBound(WireCrossSections)
        section = WireCrossSections(k)

        ' Проверяем, содержит ли строка определенные подстроки
        If InStr(section, "0.75") > 0 Or InStr(section, "1.5") > 0 Then
            If UBound(A) = 0 And A(0) = "" Then
                A(0) = section
            Else
                ReDim Preserve A(UBound(A) + 1)
                A(UBound(A)) = section
            End If
		ElseIf InStr(section, "10") > 0 Or InStr(section, "16") > 0 Or InStr(section, "25") > 0 Or InStr(section, "35") > 0 Then
            If UBound(C) = 0 And C(0) = "" Then
                C(0) = section
            Else
                ReDim Preserve C(UBound(C) + 1)
                C(UBound(C)) = section
            End If	
        ElseIf InStr(section, "2.5") > 0 Or InStr(section, "4") > 0 Or InStr(section, "6") > 0 Then
            If UBound(B) = 0 And B(0) = "" Then
                B(0) = section
            Else
                ReDim Preserve B(UBound(B) + 1)
                B(UBound(B)) = section
            End If
        
        ElseIf InStr(section, "50") > 0 Or InStr(section, "70") > 0 Then
            If UBound(E) = 0 And E(0) = "" Then
                E(0) = section
            Else
                ReDim Preserve E(UBound(E) + 1)
                E(UBound(E)) = section
            End If
        End If
    Next

    ' Выводим результаты
    App.PutInfo 0, "Длина трубок:"

    ' Проверяем массив A
    If UBound(A) >= 0 And A(0) <> "" Then
        tubeALength = (UBound(A) + 1) * 2 * 0.015 ' Длина трубки для массива A
		App.PutInfo 0, "сечение 0,75-1,5"
        App.PutInfo 0, "Трубка термоусаживаемая для термотрансферной печати ТМАРК-НГ-2П-3,2/1,6Б (50М) + риббон ТУ 22.21.29-005-65321637-2019 арт. КНГ2П-032-Б50"
		App.PutInfo 0, tubeALength
    End If

    ' Проверяем массив B
    If UBound(B) >= 0 And B(0) <> "" Then
        tubeBLength = (UBound(B) + 1) * 2 * 0.015 ' Длина трубки для массива B
		App.PutInfo 0, "сечение 2,5-6"
        App.PutInfo 0, "Трубка термоусаживаемая для термотрансферной печати ТМАРК-НГ-2П-6,4/3,2Б (50М) + риббон ТУ 22.21.29-005-65321637-2019 арт. КНГ2П-064-Б50"
		App.PutInfo 0, tubeBLength
    End If

    ' Проверяем массив C
    If UBound(C) >= 0 And C(0) <> "" Then
        tubeCLength = (UBound(C) + 1) * 2 * 0.015 ' Длина трубки для массива C
		App.PutInfo 0, "сечение 10-35"
        App.PutInfo 0, "Трубка термоусаживаемая для термотрансферной печати ТМАРК-НГ-2П-12,7/6,4Б (50М) + риббон ТУ 22.21.29-005-65321637-2019 арт. КНГ2П-127-Б50"
		App.PutInfo 0, tubeCLength
    End If

    ' Проверяем массив E
    If UBound(E) >= 0 And E(0) <> "" Then
        tubeELength = (UBound(E) + 1) * 2 * 0.015 ' Длина трубки для массива E
		App.PutInfo 0, "сечение 50-70"
        App.PutInfo 0, "Трубка термоусаживаемая для термотрансферной печати ТМАРК-НГ-2П-19,1/9,5Б (50М) + риббон ТУ 22.21.29-005-65321637-2019 арт. КНГ2П-191-Б50"
		App.PutInfo 0, tubeELength
    End If
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