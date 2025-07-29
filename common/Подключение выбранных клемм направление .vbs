' Скрипт переподключает провода на нижние зажимы клемм

Set e3Application = CreateObject( "CT.Application" ) 
Set job = e3Application.CreateJobObject()
Set symbol = job.CreateSymbolObject()
Set pin = job.CreatePinObject()
Set devicePin = job.CreatePinObject()
Set device = job.CreateDeviceObject()

'Dim connectionDirection : connectionDirection = "7"  'снизу 

' 0 - Подключается с любого направления
' 1 - Подключается справа
' 3 - Подключается сверху
' 5 - Подключается слева
' 7 - Подключается снизу
' 9 - Вертикальный: Подключается сверху или снизу
' 10 - Горизонтальный: Подключается слева или справа
' 14 - Автоматический: Направление подключения зависит от положения штифта

connectionDirection = InputBox("0 - любое" & Chr(13)& "1 - справа "& Chr(13)& _
							"3 - сверху" & Chr(13)& "5 - слева "& Chr(13)& _
							"7- снизу" & Chr(13)& "9 - вертикально"& Chr(13)& _
							"10- горизонтально" & Chr(13)& "14 - автоматически"& Chr(13)_
							,"Подключение проводов: ", "")

 

termCount = job.GetSelectedTerminalIds( terminalIds ) 'выбрать клеммы


	If termCount > 0 Then 
		e3Application.PutInfo 0, "Выбрано клемм:" & termCount  

		For terminalIndex = 1 To termCount  'перебор массива выбранных клемм
		terminalId = device.SetId( terminalIds( terminalIndex )  )

		terminalName = device.GetName()

		result = device.GetPinIds( pinIds )					'получить пины устройства

			If result > 0 Then							'если есть пины, то
				For pinIndex = 1 To result 				'перебор пинов
				pinId = pin.SetId( pinIds( pinIndex ) ) 'устанавливается активный пин
				pinName = pin.GetName()					'получить имя пина


					result2 = pin.SetPhysicalConnectionDirection( connectionDirection )  'устанавливается направление подключения
				Select Case result2
					Case 0
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено на любое"
					Case 1
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено на правое"
					Case 3
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено верхнее"
					Case 5
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено на левое"
					Case 7
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено нижнее"
					Case 9
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено на вертикальное (сверху-снизу)"
					Case 10
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено горизонтальное (справа-слева)"
					Case 14
						message = "Клеммник " & terminalName & " ; клемма " & pinName & " направление подключения изменено автоматическое"
					End Select
				e3Application.PutInfo 0, message        'output result of operation

				Next
			End If
		Next
	End If


Set device = Nothing
Set devicePin = Nothing
Set pin = Nothing
Set symbol = Nothing
Set job = Nothing 
Set e3Application = Nothing 
