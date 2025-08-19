'*******************************************************************************
' �������� �������: E3_FindPinLocation_FromExcelInput
' �����: E3.series VBScript Assistant
' ����: 06.08.2025
' ��������: ������ ��� ������ ����� ���������� (F3) � ���� (G3) �� Excel.
'           ����� ����������� � �������� ������� E3.series.
'           ��� ������� ���������� ���� ������������ ��� ������� �� �����.
'           ��������� ���� ����������� ����� � ������� Excel, ������� �� ������ 3.
'*******************************************************************************

Option Explicit

' --- ��������� ---
Const EXCEL_SHEET_NAME = "����1"         ' ������� ��� ����� � Excel, ���� ��� ����������
Const START_DATA_ROW = 3               ' ��������� ������ ��� ������ ������ � Excel
' >>> ���� � ����� Excel �� ��������� <<<
Const EXCEL_FILE_PATH_DEFAULT = "C:\Users\SEK\Desktop\DWG_4_E3\����� �����\�������\���\�������\��1\������ ��������� � ������_��1.xlsx"
' >>> ��������� ��� ������� ��������� <<<
Const FRAGMENT_PATH = "C:\Users\SEK\Desktop\DWG_4_E3\����_����\terminal2.e3p"
Const Y_OFFSET = 150                    ' �������� �� Y �� ��������� ���� (�������� �� Y)

' --- ������� ������������ ---
Call Main()

Sub Main()
    ' --- ������������� �������� E3.series ---
    Dim e3App, job, device, pin, sheet
    Dim deviceName, pinName
    Dim deviceId, pinId
    Dim EXCEL_FILE_PATH
    
    On Error Resume Next
    ' ������� �������� ��� ���������� ��������� E3.series
    Set e3App = GetObject(, "CT.Application")
    
    If e3App Is Nothing Then
        ' ���� E3.series �� �������, �������� ������� ����� ���������
        Set e3App = CreateObject("CT.Application")
        If e3App Is Nothing Then
            MsgBox "E3.series Application �� ������� ��� �� ������.", vbCritical, "������ E3.series"
            Exit Sub
        End If
    End If
    On Error GoTo 0
    
    ' ������� ������� E3.series
    Set job = e3App.CreateJobObject()
    
    ' ��������, ��� job ������ ������� ������
    On Error Resume Next
    Set device = job.CreateDeviceObject()
    If device Is Nothing Then
        e3App.PutInfo 2, "�� ������� ������� ������ Device. ���������, ��� ������ E3.series ������."
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    On Error GoTo 0
    
    Set pin = job.CreatePinObject()
    Set sheet = job.CreateSheetObject() ' ��������� ������ ��� ������ � ������� �����
    
    e3App.PutInfo 0, "������ �������: ����� ������� ����� �� ����� �� Excel."
    
    ' --- ����������� ���� � Excel ����� ---
    EXCEL_FILE_PATH = EXCEL_FILE_PATH_DEFAULT
    
    ' ���������, ���������� �� ���� �� ���� �� ���������
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(EXCEL_FILE_PATH) Then
        e3App.PutInfo 1, "���� Excel �� ���� �� ��������� '" & EXCEL_FILE_PATH & "' �� ������."
        EXCEL_FILE_PATH = InputBox("���� Excel �� ���� �� ��������� �� ������. ����������, ������� ������ ���� � ������ Excel �����:", "���� � Excel �����", "C:\Temp\�������.xlsx")
    End If
    Set fso = Nothing

    If Trim(EXCEL_FILE_PATH) = "" Then
        e3App.PutInfo 2, "���� � Excel ����� �� ��� ������. ������ �������."
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    
    ' --- ������������� �������� Excel ---
    Dim objExcel, objWorkbook, objSheet
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If objExcel Is Nothing Then
        e3App.PutInfo 2, "�� ������� ��������� ���������� Excel. ���������, ��� Excel ����������."
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    objExcel.Visible = False
    
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_FILE_PATH)
    If objWorkbook Is Nothing Then
        e3App.PutInfo 2, "�� ������� ������� ���� Excel: " & EXCEL_FILE_PATH & ". ���������, ��� ���� ���������� � �� �����."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    
    Set objSheet = objWorkbook.Sheets(EXCEL_SHEET_NAME)
    If objSheet Is Nothing Then
        e3App.PutInfo 2, "�� ������� ����� ���� '" & EXCEL_SHEET_NAME & "' � ����� Excel. ��������� ��� �����."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If
    On Error GoTo 0
    
    ' --- ���������� ��������� ����������� ������ � ����� Excel ---
    Dim lastRow
    On Error Resume Next
    lastRow = objSheet.Cells.SpecialCells(11).Row ' xlCellTypeLastCell = 11
    If Err.Number <> 0 Then
        e3App.PutInfo 1, "�� ������� ���������� ��������� ������ � Excel. ��������, ���� ����. ������: " & Err.Description
        lastRow = START_DATA_ROW
        Err.Clear
    End If
    On Error GoTo 0

    If lastRow < START_DATA_ROW Then
        e3App.PutInfo 1, "� Excel ����� ��� ������, ������� �� ������ " & START_DATA_ROW & ". ������ ��������."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, sheet, e3App)
        Exit Sub
    End If

    ' --- ��������� ���� �� ���� ����������� ������� ---
    Dim currentRow
    For currentRow = START_DATA_ROW To lastRow
        e3App.PutInfo 0, "--- ��������� ������: " & currentRow & " ---"
        
        ' --- ������ ����� ���������� � ���� �� Excel ��� ������� ������ ---
        deviceName = Trim(CStr(objSheet.Cells(currentRow, 6).Value)) ' ������� F
        pinName = Trim(CStr(objSheet.Cells(currentRow, 7).Value))    ' ������� G
        
        ' --- ������ ������� ��� ������� ��������� �� ������� H ---
        Dim fragmentCondition
        fragmentCondition = Trim(CStr(objSheet.Cells(currentRow, 8).Value)) ' ������� H
        
        ' --- ������ �������� ��� ����������� �� �������� I � J ---
        Dim masterPin1, masterPin2
        masterPin1 = Trim(CStr(objSheet.Cells(currentRow, 9).Value))  ' ������� I - ��� ������� ����������
        masterPin2 = Trim(CStr(objSheet.Cells(currentRow, 10).Value)) ' ������� J - ��� ������� ����������
        
        ' --- ��������� ������� "-XT" ����� ---
        If InStr(1, fragmentCondition, "-XT", vbTextCompare) = 0 Then
            e3App.PutInfo 0, "������ " & currentRow & ": ������� '-XT' �� ������� � ������� H ('" & fragmentCondition & "'). ������ ���������."
            ' ��������� � ��������� ������ ��� ���������� ���������
        ElseIf deviceName = "" Or pinName = "" Then
            e3App.PutInfo 1, "������ " & currentRow & ": ������� ������� '-XT', �� ��� ���������� (F" & currentRow & ") ��� ��� ���� (G" & currentRow & ") �����. ������ ���������."
        Else
            e3App.PutInfo 0, "������ " & currentRow & " ������� ������� '-XT'. ���������: ���������� = '" & deviceName & "', ��� = '" & pinName & "', MasterPin1 = '" & masterPin1 & "', MasterPin2 = '" & masterPin2 & "'"
            
            ' --- ����� ����������: �������� �� ���� ����������� ---
            e3App.PutInfo 0, "��� ����������: '" & deviceName & "' ����� ������� ��������..."
            Dim allDeviceIds, totalDeviceCount
            Dim currentDeviceId
            Dim foundDeviceId : foundDeviceId = 0

            totalDeviceCount = job.GetAllDeviceIds(allDeviceIds)

            If totalDeviceCount > 0 Then
                Dim k
                For k = 1 To totalDeviceCount
                    currentDeviceId = allDeviceIds(k)
                    device.SetId currentDeviceId
                    Dim currentDeviceName
                    currentDeviceName = device.GetName()

                    If LCase(currentDeviceName) = LCase(deviceName) Then
                        foundDeviceId = currentDeviceId
                        Exit For
                    End If
                Next
            End If

            If foundDeviceId = 0 Then
                e3App.PutInfo 1, "������ " & currentRow & ": ���������� '" & deviceName & "' �� ������� � �������."
            Else
                deviceId = foundDeviceId
                device.SetId deviceId
                e3App.PutInfo 0, "������ " & currentRow & ": ���������� '" & deviceName & "' �������. ID: " & deviceId
                
                ' --- ����� ���� �� ��������� ���������� ---
                e3App.PutInfo 0, "��� ���: '" & pinName & "' �� ���������� '" & deviceName & "' ����� ������� ��������..."
                Dim allPinIds, totalPinCount
                Dim currentPinId
                Dim foundPinId : foundPinId = 0

                totalPinCount = device.GetAllPinIds(allPinIds)

                If totalPinCount > 0 Then
                    Dim l
                    For l = 1 To totalPinCount
                        currentPinId = allPinIds(l)
                        pin.SetId currentPinId
                        Dim currentPinName
                        currentPinName = pin.GetName()

                        If LCase(currentPinName) = LCase(pinName) Then
                            foundPinId = currentPinId
                            Exit For
                        End If
                    Next
                End If

                If foundPinId = 0 Then
                    e3App.PutInfo 1, "������ " & currentRow & ": ��� '" & pinName & "' �� ������ �� ���������� '" & deviceName & "'."
                Else
                    pinId = foundPinId
                    pin.SetId pinId
                    e3App.PutInfo 0, "������ " & currentRow & ": ��� '" & pinName & "' ������ �� ���������� '" & deviceName & "'. ID ����: " & pinId
                    
                    ' --- ��������� ������� ���� �� ����� ---
                    e3App.PutInfo 0, "������ " & currentRow & ": ����������� ������� ���� '" & pinName & "' �� �����..."
                    
                    Dim xPosition, yPosition, gridDescription, columnValue, rowValue
                    Dim result
                    
                    On Error Resume Next
                    result = pin.GetSchemaLocation(xPosition, yPosition, gridDescription, columnValue, rowValue)
                    On Error GoTo 0
                    
                    If result = 0 Then
                        e3App.PutInfo 1, "������ " & currentRow & ": ��� '" & pinName & "' (" & pinId & ") �� �������� �� ����� ��� ��������� ������."
                    Else
                        ' ������������� ID ����� ����� ��� ��������� ��� �����
                        sheet.SetId result
                        Dim sheetName
                        sheetName = sheet.GetName()
                        
                        e3App.PutInfo 0, "������ " & currentRow & ": ������� ���� '" & pinName & "' (" & pinId & ") �� �����:"
                        e3App.PutInfo 0, "    ���� �����: " & sheetName & " (ID: " & result & ")"
                        e3App.PutInfo 0, "    X ����������: " & xPosition
                        e3App.PutInfo 0, "    Y ����������: " & yPosition
                        e3App.PutInfo 0, "    �����: " & gridDescription
                        e3App.PutInfo 0, "    �������: " & columnValue
                        e3App.PutInfo 0, "    ������: " & rowValue
                        
                        ' --- ������� ��������� �� ����������� ���� (������� "-XT" ��� ���������) ---
                        Dim fragmentX, fragmentY
                        fragmentX = xPosition
                        fragmentY = yPosition - Y_OFFSET ' �������� �������� �� Y ����������
                        
                        e3App.PutInfo 0, "������ " & currentRow & ": ������� ��������� terminal2.e3p �� ����������� X=" & fragmentX & ", Y=" & fragmentY
                        
                        Dim fragmentResult
                        fragmentResult = PlaceFragmentOnSheet(sheet, result, FRAGMENT_PATH, "", fragmentX, fragmentY, e3App, currentRow)
                        
                        ' --- �������������� ��������� -sXT1 ����� �������� ������� ��������� ---
                        If fragmentResult = 0 Then
                            e3App.PutInfo 0, "������ " & currentRow & ": �������� ������� ��������. ������ �������������� ��������� -sXT1..."
                            Call RenameDevicesAfterFragment(job, e3App, currentRow, fragmentCondition, masterPin1, masterPin2)
                        Else
                            e3App.PutInfo 1, "������ " & currentRow & ": �������� �� �������� (���: " & fragmentResult & "). �������������� �� �����������."
                        End If
                    End If
                End If
            End If
        End If
    Next

    ' --- ������������ ������� Excel �������� ---
    Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
    
    ' --- ������� ��������� � ���������� ������� ---
    e3App.PutInfo 0, "������ ��������."

    ' --- ������� �������� E3.series ---
    Call CleanUpE3Objects(pin, device, job, sheet, e3App)
    
End Sub

' --- ������� ��� ������� ��������� �� ���� ����� ---
Function PlaceFragmentOnSheet(sheet, sheetId, fragmentPath, version, xPosition, yPosition, e3App, rowNumber)
    On Error Resume Next
    
    ' ������������� ID �������� �����
    sheet.SetId sheetId
    Dim sheetName
    sheetName = sheet.GetName()
    
    ' ��������� ��������
    Dim result
    result = sheet.PlacePart(fragmentPath, version, xPosition, yPosition, 0.0)
    
    ' ������������ ��������� �������
    Dim message
    Select Case result
        Case 9
            message = "������ ������� ��������� �� ���� " & sheetName & " (" & sheetId & "): ������������� ������ ����� ���������"
        Case 3
            message = "������ ������� ��������� �� ���� " & sheetName & " (" & sheetId & "): �������� ��� ��������� ��� ������"
        Case 0
            message = "�������� ������� �������� �� ���� " & sheetName & " (" & sheetId & ") �� ����������� X=" & xPosition & ", Y=" & yPosition
        Case -1
            message = "������ ������� ��������� �� ���� " & sheetName & " (" & sheetId & "): �������� ������� �� ���������� ������ � ����������� ��������� '������������ ������� �����'"
        Case -2
            message = "������ ������� ��������� �� ���� " & sheetName & " (" & sheetId & "): �������� �������� ����� � �� ����������� ��������� '������������ ������� �����'"
        Case -3
            message = "������ ������� ��������� �� ���� " & sheetName & " (" & sheetId & "): �������� ��� �������� ��� ������ ������� ��������� �� ����������� X=" & xPosition & ", Y=" & yPosition
        Case -4
            message = "������ ������� ��������� �� ���� " & sheetName & " (" & sheetId & "): ���� ������������"
        Case Else
            message = "������ ������� ��������� �� ���� " & sheetName & " (" & sheetId & "): ��� ������ " & result
    End Select
    
    e3App.PutInfo 0, "������ " & rowNumber & ": " & message
    
    On Error GoTo 0
    PlaceFragmentOnSheet = result
End Function

' --- ������� �������������� ��������� -sXT1 � -XT666 ����� ������� ��������� ---
Sub RenameDevicesAfterFragment(job, e3App, rowNumber, newDeviceName, masterPin1Value, masterPin2Value)
    On Error Resume Next
    
    Dim renameDevice
    Set renameDevice = job.CreateDeviceObject()
    
    Dim deviceIds
    Dim result
    result = job.GetAllDeviceIds(deviceIds)
    
    Dim foundCount
    foundCount = 0
    
    e3App.PutInfo 0, "������ " & rowNumber & ": === ����� ��������� -sXT1 ==="
    
    If result > 0 Then
        e3App.PutInfo 0, "������ " & rowNumber & ": ����� ��������� � �������: " & result
        
        ' ������� ������ ��� ���������� -sXT1 � �������� �� ����������
        Dim sxtDevices()
        Dim sxtCount
        sxtCount = 0
        
        Dim i, name, currentMasterPin
        For i = 1 To result
            renameDevice.SetId deviceIds(i)
            name = renameDevice.GetName()
            
            If name = "-sXT1" Then
                ' ���������, ���� �� � ���������� ���������
                currentMasterPin = renameDevice.GetMasterPinName()
                
                e3App.PutInfo 0, "������ " & rowNumber & ": --- ������� ���������� -sXT1 ---"
                e3App.PutInfo 0, "������ " & rowNumber & ": ID ����������: " & deviceIds(i)
                e3App.PutInfo 0, "������ " & rowNumber & ": ������� ���������: '" & currentMasterPin & "'"
                
                ' ���� ��������� ���������� (�� ������)
                If Len(Trim(currentMasterPin)) > 0 Then
                    sxtCount = sxtCount + 1
                    ReDim Preserve sxtDevices(sxtCount - 1)
                    sxtDevices(sxtCount - 1) = deviceIds(i)
                    e3App.PutInfo 0, "������ " & rowNumber & ": ���������� ��������� � ������ ��� ��������� (�" & sxtCount & ")"
                Else
                    e3App.PutInfo 0, "������ " & rowNumber & ": ���������� ��������� - ��� ����������"
                End If
            End If
        Next
        
        e3App.PutInfo 0, "������ " & rowNumber & ": === ��������� ��������� ��������� ==="
        e3App.PutInfo 0, "������ " & rowNumber & ": ��������� � ������������ ��� ���������: " & sxtCount
        
        ' ������ ������������ ��������� ���������� � ������������
        Dim j, newName, resultSet, resultPin, finalMasterPin
        For j = 0 To sxtCount - 1
            If j >= 2 Then
                e3App.PutInfo 0, "������ " & rowNumber & ": ������� ����� ���� ��������� � ������������. ��������� �� ����������."
                Exit For
            End If
            
            renameDevice.SetId sxtDevices(j)
            foundCount = j + 1
            
            e3App.PutInfo 0, "������ " & rowNumber & ": --- ��������� ���������� #" & foundCount & " ---"
            
            ' �������������� ���������� � �������� �� ������� H
            If Len(Trim(newDeviceName)) > 0 Then
                resultSet = renameDevice.SetName(newDeviceName)
                
                If resultSet = 0 Then
                    e3App.PutInfo 0, "������ " & rowNumber & ": ������ ��� �������������� ���������� #" & foundCount & " � '" & newDeviceName & "'"
                Else
                    e3App.PutInfo 0, "������ " & rowNumber & ": ���������� #" & foundCount & " ������������� � '" & newDeviceName & "'"
                End If
            Else
                e3App.PutInfo 1, "������ " & rowNumber & ": ����� ��� ���������� (������� H) ������. ���������� #" & foundCount & " �� �������������."
            End If
            
            ' ��������� ����������
            If foundCount = 1 Then
                ' ���������� �������� �� ������� I ��� ������� ����������
                If Len(Trim(masterPin1Value)) > 0 Then
                    resultPin = renameDevice.SetMasterPinName(masterPin1Value)
                    If resultPin = 0 Then
                        e3App.PutInfo 0, "������ " & rowNumber & ": ������ ��� ��������� ���������� '" & masterPin1Value & "' ��� ���������� #" & foundCount
                    Else
                        e3App.PutInfo 0, "������ " & rowNumber & ": ��������� ���������� #" & foundCount & " ���������� �: " & masterPin1Value
                    End If
                Else
                    e3App.PutInfo 1, "������ " & rowNumber & ": �������� ���������� ��� ������� ���������� (������� I) ������. ��������� �� �������."
                End If
            ElseIf foundCount = 2 Then
                ' ���������� �������� �� ������� J ��� ������� ����������
                If Len(Trim(masterPin2Value)) > 0 Then
                    resultPin = renameDevice.SetMasterPinName(masterPin2Value)
                    If resultPin = 0 Then
                        e3App.PutInfo 0, "������ " & rowNumber & ": ������ ��� ��������� ���������� '" & masterPin2Value & "' ��� ���������� #" & foundCount
                    Else
                        e3App.PutInfo 0, "������ " & rowNumber & ": ��������� ���������� #" & foundCount & " ���������� �: " & masterPin2Value
                    End If
                Else
                    e3App.PutInfo 1, "������ " & rowNumber & ": �������� ���������� ��� ������� ���������� (������� J) ������. ��������� �� �������."
                End If
            End If
            
            ' �������� ����������
            finalMasterPin = renameDevice.GetMasterPinName()
            e3App.PutInfo 0, "������ " & rowNumber & ": �������� ���������: '" & finalMasterPin & "'"
        Next
        
        If sxtCount = 0 Then
            e3App.PutInfo 0, "������ " & rowNumber & ": �� ������� �� ������ ���������� -sXT1 � �����������."
        End If
        
    Else
        e3App.PutInfo 0, "������ " & rowNumber & ": ������: ���������� � ������� �� �������."
    End If
    
    e3App.PutInfo 0, "������ " & rowNumber & ": === �������������� ��������� ==="
    
    ' �������
    Set renameDevice = Nothing
    On Error GoTo 0
End Sub

' --- ��������������� ������������ ��� ������� �������� ---
Sub CleanUpExcelObjects(ByRef objSheet, ByRef objWorkbook, ByRef objExcel)
    On Error Resume Next
    If Not objWorkbook Is Nothing Then objWorkbook.Close False
    If Not objExcel Is Nothing Then objExcel.Quit
    Set objSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0
End Sub

Sub CleanUpE3Objects(ByRef pin, ByRef device, ByRef job, ByRef sheet, ByRef e3App)
    On Error Resume Next
    Set pin = Nothing
    Set device = Nothing
    Set job = Nothing
    Set sheet = Nothing
    Set e3App = Nothing
    On Error GoTo 0
End Sub