'===============================================================================
' ������ ��������������� ���������� ����� �� ���� � E3.SERIES
'===============================================================================
' ��������: ������������� ������� ��� ���� � ������� e3.series, �������������
'           � ��������� ���� (signal name), ��������� �� � ������� ����� ����
'           ���������������� ���������� ���������.
'
' ������� ���������:
'   - wiregroupName     : ��� ������ ��������
'   - databaseWireName  : ��� ������� � ���� ������
'   - wireName          : ��� ������������ �������
'   - signalName        : ��� ���� ��� ������ �����
'
' �������� ������:
'   1. ����� ���� ����� � ������� �� ����� ������� (SignalName)
'   2. �������������� ����� ����� ����� ������� Connection
'   3. ���������� ��������� ����� �� ����� ���������� � ����
'   4. �������� ���������������� ���������� ����� ��������� ������
'
' �����: ������
' ����:  [���� ��������]
'===============================================================================
Option Explicit

' --- ��������� ���������� �� HTA ---
If WScript.Arguments.Count <> 4 Then
    MsgBox "��������� 4 ���������: wiregroupName, databaseWireName, wireName, signalName", vbCritical, "������ �������"
    WScript.Quit 1
End If

Dim wiregroupName, databaseWireName, wireName, signalName
wiregroupName     = WScript.Arguments(0)
databaseWireName  = WScript.Arguments(1)
wireName          = WScript.Arguments(2)
signalName        = WScript.Arguments(3)

' --- ������������� E3 API ---
Dim app, job, symbol, pin, device, connection
Set app        = CreateObject("CT.Application")
Set job        = app.CreateJobObject()
Set symbol     = job.CreateSymbolObject()
Set pin        = job.CreatePinObject()
Set device     = job.CreateDeviceObject()
Set connection = job.CreateConnectionObject()

Dim dictPinIds
Set dictPinIds = CreateObject("Scripting.Dictionary")

app.PutInfo 0, "�������� ����� ����� ����: " & signalName

Call FIND_ALL_PINS_BY_SIGNAL()
Call FIND_PINS_BY_CONNECTIONS()
Call CREATE_CONNECTIONS()

app.PutInfo 0, "��������� ���������."

' === ����� ����� �� SignalName ===
Sub FIND_ALL_PINS_BY_SIGNAL()
    app.PutInfo 0, "����� ���� ����� � ������� �� signal name..."

    Dim deviceIds, pinIds
    Dim deviceCount, pinCount, deviceIndex, pinIndex, pinId, deviceId
    deviceCount = job.GetDeviceIds(deviceIds)
    app.PutInfo 0, "������� ��������� � �������: " & deviceCount

    For deviceIndex = 1 To deviceCount
        deviceId = device.SetId(deviceIds(deviceIndex))
        Dim deviceName
        deviceName = device.GetName()

        pinCount = device.GetPinIds(pinIds)
        If pinCount > 0 Then
            For pinIndex = 1 To pinCount
                pinId = pin.SetId(pinIds(pinIndex))
                Dim pinName, pinSignalName
                pinName = pin.GetName()
                pinSignalName = pin.GetSignalName()

                If pinSignalName = signalName Then
                    If Not dictPinIds.Exists(pinId) Then
                        dictPinIds.Add pinId, deviceName & "." & pinName
                        app.PutInfo 0, "������ ��� (signal): " & deviceName & "." & pinName & " (ID: " & pinId & ")"
                    End If
                End If
            Next
        End If
    Next

    app.PutInfo 0, "������� ����� �� signal name: " & dictPinIds.Count
End Sub

' === ����� ����� connections ===
Sub FIND_PINS_BY_CONNECTIONS()
    app.PutInfo 0, "����� ����� ����� connections..."

    Dim connectionIds, pinIds, connectionIndex, connectionId
    Dim connectedPinIds
    Set connectedPinIds = CreateObject("Scripting.Dictionary")

    Dim connectionCount
    connectionCount = job.GetConnectionIds(connectionIds)
    app.PutInfo 0, "������� connections � �������: " & connectionCount

    For connectionIndex = 1 To connectionCount
        connectionId = connection.SetId(connectionIds(connectionIndex))
        If connection.GetSignalName() = signalName Then
            Dim pinCount, pinIndex, pinId
            pinCount = connection.GetPinIds(pinIds)
            For pinIndex = 1 To pinCount
                pinId = pinIds(pinIndex)
                connectedPinIds(pinId) = True
            Next
            app.PutInfo 0, "Connection ID: " & connectionId & " �������� " & pinCount & " ����� ���� " & signalName
        End If
    Next

    Dim connectedPinKeys, key
    connectedPinKeys = connectedPinIds.Keys()
    For Each key In connectedPinKeys
        If Not dictPinIds.Exists(key) Then
            pin.SetId key
            Dim pinName, devName
            pinName = pin.GetName()
            device.SetId key
            devName = device.GetName()
            dictPinIds.Add key, devName & "." & pinName
            app.PutInfo 0, "������ ��� (connection): " & devName & "." & pinName & " (ID: " & key & ")"
        End If
    Next

    app.PutInfo 0, "����� ������� ���������� �����: " & dictPinIds.Count
End Sub

' === �������� ���������� ===
Sub CREATE_CONNECTIONS()
    If dictPinIds.Count < 2 Then
        app.PutInfo 0, "������������ ����� ��� �������� ���������� (�������: " & dictPinIds.Count & ")"
        Exit Sub
    End If

    app.PutInfo 0, "�������� ���������� ����� ������..."
    Call SORT_FOUND_PINS()

    Dim pinKeys, i, firstPinId, secondPinId
    pinKeys = dictPinIds.Keys()

    For i = 0 To dictPinIds.Count - 2
        firstPinId  = pinKeys(i)
        secondPinId = pinKeys(i + 1)

        app.PutInfo 0, "���������: " & dictPinIds(firstPinId) & " -> " & dictPinIds(secondPinId)
        If CREATE_WIRE(firstPinId, secondPinId) Then
            app.PutInfo 0, "���������� ������� �������"
        Else
            app.PutError 0, "������ �������� ����������"
        End If
    Next

    app.PutInfo 0, "������� ����������: " & (dictPinIds.Count - 1)
End Sub

' === ���������� ��������� ����� ===
Sub SORT_FOUND_PINS()
    If dictPinIds.Count <= 1 Then Exit Sub

    app.PutInfo 0, "���������� ��������� �����..."

    Dim pinKeys, pinCount, sortArray(), i
    pinKeys = dictPinIds.Keys()
    pinCount = dictPinIds.Count
    ReDim sortArray(pinCount - 1)

    For i = 0 To pinCount - 1
        Dim fullName, deviceName, pinName, dotPos, pinId
        pinId = pinKeys(i)
        fullName = CStr(dictPinIds(pinId))
        dotPos = InStr(fullName, ".")

        If dotPos > 0 Then
            deviceName = Left(fullName, dotPos - 1)
            pinName = Mid(fullName, dotPos + 1)
        Else
            deviceName = fullName
            pinName = ""
        End If

        sortArray(i) = Array(deviceName, pinName, pinId)
    Next

    Dim options(1, 1)
    options(0, 0) = 0 ' deviceName
    options(0, 1) = 2
    options(1, 0) = 1 ' pinName
    options(1, 1) = 2

    app.SortArrayByIndexEx sortArray, options

    dictPinIds.RemoveAll

    For i = 0 To pinCount - 1
        deviceName = sortArray(i)(0)
        pinName = sortArray(i)(1)
        pinId = sortArray(i)(2)
        dictPinIds.Add pinId, deviceName & "." & pinName
        app.PutInfo 0, (i + 1) & ". " & deviceName & "." & pinName & " (ID: " & pinId & ")"
    Next
End Sub

' === �������� ������� ����� ����� ������ ===
Function CREATE_WIRE(firstPinId, secondPinId)
    CREATE_WIRE = False

    If firstPinId <= 0 Or secondPinId <= 0 Then
        app.PutError 0, "�������� ID �����: " & firstPinId & ", " & secondPinId
        Exit Function
    End If

    Dim cableCount, cableIds, cableId, cableName
    cableCount = job.GetCableIds(cableIds)

    If cableCount = 0 Then
        app.PutError 0, "�� ������� ������ � �������"
        Exit Function
    End If

    Dim i, result, actualWireName
    For i = 1 To cableCount
        cableId = device.SetId(cableIds(i))
        If device.IsWireGroup() = 1 Then
            result = pin.CreateWire(wireName, wiregroupName, databaseWireName, cableId, 0, 0)
            If result > 0 Then
                pin.SetEndPinId 1, firstPinId
                pin.SetEndPinId 2, secondPinId
                actualWireName = pin.GetName()
                app.PutInfo 0, "������ ������: " & actualWireName & " (" & wiregroupName & ", " & databaseWireName & ")"
                CREATE_WIRE = True
                Exit Function
            Else
                cableName = device.GetName()
                app.PutError 0, "������ �������� ������� � ������: " & cableName
            End If
        End If
    Next

    app.PutError 0, "�� ������ ���������� wire group ��� �������� �������"
End Function

' === ������� ===
Set dictPinIds = Nothing
Set connection = Nothing
Set device     = Nothing
Set pin        = Nothing
Set symbol     = Nothing
Set job        = Nothing
Set app        = Nothing
