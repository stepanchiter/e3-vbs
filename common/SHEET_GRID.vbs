Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject
Set Cab = Job.CreateDeviceObject
Set Cor = Job.CreatePinObject
Set Sig = Job.CreateSignalObject

If Job.GetId = 0 Then
    App.PutInfo 1, "No project opened, exiting..."
    WScript.Quit
End If


JobName = Job.GetName
App.PutInfo 0, "---- начало работы ----"



'===============================================================
' 1-й этап: переименование цепей, начинающихся с #, в формат #<.SHEET><.GRID>
'===============================================================
App.PutInfo 0, "1-й этап: переименование цепей в формат #<.SHEET><.GRID>"

' Получаем все цепи в проекте
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


App.PutInfo 0, "---- конец работы ----"