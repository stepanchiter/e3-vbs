Set e3Application = CreateObject( "CT.Application" ) 
Set job = e3Application.CreateJobObject()
Set device = job.CreateDeviceObject()
'Задаем поз.обозначение 
Dim deviceDesignation : deviceDesignation = InputBox("Новое поз.обозначение", "")

'deviceCount = job.GetTreeSelectedAllDeviceIds( deviceIds )        

termCount = job.GetTreeSelectedTerminalIds( terminalIds ) 'выбрать клеммы

If termCount > 0 Then 
		e3Application.PutInfo 0, "Выбрано клемм:" & termCount  

		For terminalIndex = 1 To termCount  'перебор массива выбранных клемм
		terminalId = device.SetId( terminalIds( terminalIndex )  )
		terminalName = device.GetName()
		
		'deviceId = device.SetId( deviceIds( 1 ) )        'первое устройство в перечне 
		deviceName = device.GetName()

		result = device.SetName( deviceDesignation )

		If result = 0 Then
			message = "Устройство " & deviceId & ": Ошибка поз. обозначения" 
		Else
			message = "Устройство " & deviceId & ": поз. обозначение изменено с " & deviceName & " на " & deviceDesignation
		End If        
		e3Application.PutInfo 0, message        'вывод результата
	Next

End If

Set device = Nothing
Set job = Nothing 
Set e3Application = Nothing 
