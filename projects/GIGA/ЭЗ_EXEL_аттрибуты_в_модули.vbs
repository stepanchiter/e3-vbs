'*******************************************************************************
' �������� �������: E3_FindDeviceAndPin_FromExcelInput_BRS_Approach
' �����: E3.series VBScript Assistant
' ����: 16.07.2025
' ��������: ������ ��� ������ ����� ���������� (F3) � ���� (G3) �� Excel.
'           ����� ����������� � �������� ������� E3.series.
'           ���������� ������ ����������� � E3.series, ����������� �������� �� ���,
'           ��� ������ ������ GetActiveJobId().
'           ��������: ����� ���������� ������ �������������� ����� �������� �� ����
'           ����������� ������� � ������������� ����� ��� ����� ��������.
'           �����, ����� ���� �������������� ����� �������� �� ���� �����
'           ���������� ���������� � ������������� ����� ��� ����� ��������.
'           ����������: ��������� � ���������� ������� ������ ��������� �� ������� �������� E3.series.
'           ���������: ������ �������� �� ����� Excel (A3-E3) � �������� ���������� ����.
'           ���������: ��������� ���� ����������� ����� � ������� Excel, ������� �� ������ 3.
'           ���������: ���� � Excel ����� �������� � ���������� �������.
'*******************************************************************************

Option Explicit

' --- ��������� ---
Const EXCEL_SHEET_NAME = "����1"         ' ������� ��� ����� � Excel, ���� ��� ����������
Const START_DATA_ROW = 3               ' ��������� ������ ��� ������ ������ � Excel
' >>> ��������� ����� ��������� ��� ���� � ����� Excel <<<
Const EXCEL_FILE_PATH_DEFAULT = "C:\Users\SEK\Desktop\DWG_4_E3\����� �����\�������\���\�������\��2\������ ��������� � ������_��2.xlsx" ' ��� ���� � ����� Excel �� ���������

' --- ������� ������������ ---
Call Main()

Sub Main()
    ' --- ������������� �������� E3.series ---
    Dim e3App, job, device, pin
    Dim deviceName, pinName
    Dim deviceId, pinId
    Dim EXCEL_FILE_PATH ' ������ ��� ����������, � �� ���������, ����� ����� ���� ������������
    
    ' ���������� ��� ������ ��������� �� Excel
    Dim tagPosition, tagDescription, plcSignalType, plcConnectionType, plcUnit
    
    On Error Resume Next
    ' ������� �������� ��� ���������� ��������� E3.series
    Set e3App = GetObject(, "CT.Application")
    
    If e3App Is Nothing Then
        ' ���� E3.series �� �������, �������� ������� ����� ���������
        Set e3App = CreateObject("CT.Application")
        If e3App Is Nothing Then
            MsgBox "E3.series Application �� ������� ��� �� ������.", vbCritical, "������ E3.series"
            Exit Sub ' ����� �� ������������
        End If
    End If
    On Error GoTo 0 ' ��������� ��������� ������ ����� ������������� e3App
    
    ' ������� ������ job. � ���� ������� ��������������, ��� �� ����� ��������
    ' � �������� ��������, ���� �� ������.
    Set job = e3App.CreateJobObject()
    
    ' ��������, ��� job ������ ������� ������ � ������ ������ (��������)
    ' ���� job.CreateDeviceObject() �� ���������, ��� ������ �� �������� � ��������.
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
    
    e3App.PutInfo 0, "������ �������: ����� ���������� � ���� �� Excel (������ ���)."
    
    ' --- ����������� ���� � Excel ����� ---
    EXCEL_FILE_PATH = EXCEL_FILE_PATH_DEFAULT ' �� ��������� ���������� ���� �� ��������
    
    ' ���������, ���������� �� ���� �� ���� �� ���������
    Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(EXCEL_FILE_PATH) Then
        e3App.PutInfo 1, "���� Excel �� ���� �� ��������� '" & EXCEL_FILE_PATH & "' �� ������."
        ' ���� ����� ���, ����������� ���� � ������������
        EXCEL_FILE_PATH = InputBox("���� Excel �� ���� �� ��������� �� ������. ����������, ������� ������ ���� � ������ Excel �����:", "���� � Excel �����", "C:\Temp\�������.xlsx")
    End If
    Set fso = Nothing

    If Trim(EXCEL_FILE_PATH) = "" Then
        e3App.PutInfo 2, "���� � Excel ����� �� ��� ������. ������ �������."
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' ����� �� ������������
    End If
    
    ' --- ������������� �������� Excel ---
    Dim objExcel, objWorkbook, objSheet
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If objExcel Is Nothing Then
        e3App.PutInfo 2, "�� ������� ��������� ���������� Excel. ���������, ��� Excel ����������."
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' ����� �� ������������
    End If
    objExcel.Visible = False ' ������ Excel
    
    Set objWorkbook = objExcel.Workbooks.Open(EXCEL_FILE_PATH)
    If objWorkbook Is Nothing Then
        e3App.PutInfo 2, "�� ������� ������� ���� Excel: " & EXCEL_FILE_PATH & ". ���������, ��� ���� ���������� � �� �����."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' ����� �� ������������
    End If
    
    Set objSheet = objWorkbook.Sheets(EXCEL_SHEET_NAME)
    If objSheet Is Nothing Then
        e3App.PutInfo 2, "�� ������� ����� ���� '" & EXCEL_SHEET_NAME & "' � ����� Excel. ��������� ��� �����."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub ' ����� �� ������������
    End If
    On Error GoTo 0 ' ��������� ��������� ������ ����� ��������� �������� Excel
    
    ' --- ���������� ��������� ����������� ������ � ����� Excel ---
    Dim lastRow
    On Error Resume Next
    ' ������������� SpecialCells(xlCellTypeLastCell) ��� ���������� ��������� ������ � �������
    lastRow = objSheet.Cells.SpecialCells(11).Row ' xlCellTypeLastCell = 11
    If Err.Number <> 0 Then
        e3App.PutInfo 1, "�� ������� ���������� ��������� ������ � Excel. ��������, ���� ����. ������: " & Err.Description
        lastRow = START_DATA_ROW ' � ������ ������, ������������� �� ��������� ������
        Err.Clear
    End If
    On Error GoTo 0

    If lastRow < START_DATA_ROW Then
        e3App.PutInfo 1, "� Excel ����� ��� ������, ������� �� ������ " & START_DATA_ROW & ". ������ ��������."
        Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
        Call CleanUpE3Objects(pin, device, job, e3App)
        Exit Sub
    End If

    ' --- ��������� ���� �� ���� ����������� ������� ---
    Dim currentRow
    For currentRow = START_DATA_ROW To lastRow
        e3App.PutInfo 0, "--- ��������� ������: " & currentRow & " ---"
        
        ' --- ������ ����� ���������� � ���� �� Excel ��� ������� ������ ---
        deviceName = Trim(CStr(objSheet.Cells(currentRow, 6).Value)) ' ������� F
        pinName = Trim(CStr(objSheet.Cells(currentRow, 7).Value))    ' ������� G
        
        ' --- ������ �������� ��� ��������� �� Excel ��� ������� ������ ---
        tagPosition = Trim(CStr(objSheet.Cells(currentRow, 1).Value))  ' ������� A
        tagDescription = Trim(CStr(objSheet.Cells(currentRow, 2).Value)) ' ������� B
        plcSignalType = Trim(CStr(objSheet.Cells(currentRow, 3).Value)) ' ������� C
        plcConnectionType = Trim(CStr(objSheet.Cells(currentRow, 4).Value)) ' ������� D
        plcUnit = Trim(CStr(objSheet.Cells(currentRow, 5).Value))      ' ������� E
        
        ' --- �������� ����������� �������� � ���������� ������ ---
        If deviceName = "" Or pinName = "" Then
            e3App.PutInfo 1, "������ " & currentRow & ": ���������, ��� ��� ��� ���������� (F" & currentRow & ") ��� ��� ���� (G" & currentRow & ") �����."
            ' ���������� ���� � ��������� ������
        Else
            e3App.PutInfo 0, "������ " & currentRow & " ��������� �� Excel: ���������� = '" & deviceName & "', ��� = '" & pinName & "'"
            e3App.PutInfo 0, "�������� ��������� �� Excel: TAG �������='" & tagPosition & "', TAG ��������='" & tagDescription & "', ���-��� �������='" & plcSignalType & "', ���-��� �����������='" & plcConnectionType & "', ���-������� ���������='" & plcUnit & "'"
            
            ' --- ����� ����������: ������ �������� �� ���� ����������� ---
            e3App.PutInfo 0, "��� ����������: '" & deviceName & "' ����� ������� ��������..."
            Dim allDeviceIds, totalDeviceCount
            Dim currentDeviceId
            Dim foundDeviceId : foundDeviceId = 0 ' ���������� ��� �������� ID ���������� ����������

            totalDeviceCount = job.GetAllDeviceIds(allDeviceIds) ' �������� ��� ID ��������� � �������

            If totalDeviceCount > 0 Then
                Dim k ' ���������� ������ ���������� ��� �����, ����� �������� ���������� � 'i' � ������������ �������
                For k = 1 To totalDeviceCount
                    currentDeviceId = allDeviceIds(k)
                    device.SetId currentDeviceId ' �������������� ������ ���������� ������� ID
                    Dim currentDeviceName
                    currentDeviceName = device.GetName() ' �������� ��� �������� ����������

                    ' ���������� ����� ��� ����� ��������
                    If LCase(currentDeviceName) = LCase(deviceName) Then
                        foundDeviceId = currentDeviceId ' ���������� �������, ��������� ID
                        Exit For ' ������� �� �����, ��� ��� ����� ����������
                    End If
                Next
            End If

            If foundDeviceId = 0 Then
                e3App.PutInfo 1, "������ " & currentRow & ": ���������� '" & deviceName & "' �� ������� � �������."
            Else
                deviceId = foundDeviceId ' ����������� ��������� ID ��� ����������� �������������
                device.SetId deviceId ' �������������� ������ ���������� ��������� ID
                e3App.PutInfo 0, "������ " & currentRow & ": ���������� '" & deviceName & "' �������. ID: " & deviceId
                
                ' --- ����� ���� �� ��������� ����������: ������ �������� �� ���� ����� ���������� ---
                e3App.PutInfo 0, "��� ���: '" & pinName & "' �� ���������� '" & deviceName & "' ����� ������� ��������..."
                Dim allPinIds, totalPinCount
                Dim currentPinId
                Dim foundPinId : foundPinId = 0 ' ���������� ��� �������� ID ���������� ����

                totalPinCount = device.GetAllPinIds(allPinIds) ' �������� ��� ID ����� �� ��������� ����������

                If totalPinCount > 0 Then
                    Dim l ' ���������� ������ ���������� ��� �����
                    For l = 1 To totalPinCount
                        currentPinId = allPinIds(l)
                        pin.SetId currentPinId ' �������������� ������ ���� ������� ID
                        Dim currentPinName
                        currentPinName = pin.GetName() ' �������� ��� �������� ����

                        ' ���������� ����� ����� ��� ����� ��������
                        If LCase(currentPinName) = LCase(pinName) Then
                            foundPinId = currentPinId ' ��� ������, ��������� ID
                            Exit For ' ������� �� �����, ��� ��� ����� ����������
                        End If
                    Next
                End If

                If foundPinId = 0 Then
                    e3App.PutInfo 1, "������ " & currentRow & ": ��� '" & pinName & "' �� ������ �� ���������� '" & deviceName & "'."
                Else
                    pinId = foundPinId ' ����������� ��������� ID ��� ����������� �������������
                    pin.SetId pinId ' �������������� ������ ���� ��������� ID
                    e3App.PutInfo 0, "������ " & currentRow & ": ��� '" & pinName & "' ������ �� ���������� '" & deviceName & "'. ID ����: " & pinId
                    
                    ' --- ������ �������� � �������� ���������� ���� ---
                    e3App.PutInfo 0, "������ " & currentRow & ": ������ ��������� ��� ���� '" & pinName & "'..."
                    
                    On Error Resume Next ' �������� ��������� ������ ��� SetAttributeValue
                    
                    ' ��������� � ���������� 'TAG �������'
                    If tagPosition <> "" Then
                        If pin.SetAttributeValue("TAG �������", tagPosition) = 0 Then
                            e3App.PutInfo 1, "������ " & currentRow & ": ������ ��� ��������� �������� 'TAG �������' ��� ���� '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "������ " & currentRow & ": ������� 'TAG �������' ������� ���������� � '" & tagPosition & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "������ " & currentRow & ": �������� ��� 'TAG �������' � Excel �����, ������� �� �������."
                    End If
                    
                    ' ��������� � ���������� 'TAG ��������'
                    If tagDescription <> "" Then
                        If pin.SetAttributeValue("TAG ��������", tagDescription) = 0 Then
                            e3App.PutInfo 1, "������ " & currentRow & ": ������ ��� ��������� �������� 'TAG ��������' ��� ���� '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "������ " & currentRow & ": ������� 'TAG ��������' ������� ���������� � '" & tagDescription & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "������ " & currentRow & ": �������� ��� 'TAG ��������' � Excel �����, ������� �� �������."
                    End If

                    ' ��������� � ���������� '��� - ��� �������'
                    If plcSignalType <> "" Then
                        If pin.SetAttributeValue("��� - ��� �������", plcSignalType) = 0 Then
                            e3App.PutInfo 1, "������ " & currentRow & ": ������ ��� ��������� �������� '��� - ��� �������' ��� ���� '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "������ " & currentRow & ": ������� '��� - ��� �������' ������� ���������� � '" & plcSignalType & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "������ " & currentRow & ": �������� ��� '��� - ��� �������' � Excel �����, ������� �� �������."
                    End If

                    ' ��������� � ���������� '��� - ��� �����������'
                    If plcConnectionType <> "" Then
                        If pin.SetAttributeValue("��� - ��� �����������", plcConnectionType) = 0 Then
                            e3App.PutInfo 1, "������ " & currentRow & ": ������ ��� ��������� �������� '��� - ��� �����������' ��� ���� '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "������ " & currentRow & ": ������� '��� - ��� �����������' ������� ���������� � '" & plcConnectionType & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "������ " & currentRow & ": �������� ��� '��� - ��� �����������' � Excel �����, ������� �� �������."
                    End If

                    ' ��������� � ���������� '��� - ������� ���������'
                    If plcUnit <> "" Then
                        If pin.SetAttributeValue("��� - ������� ���������", plcUnit) = 0 Then
                            e3App.PutInfo 1, "������ " & currentRow & ": ������ ��� ��������� �������� '��� - ������� ���������' ��� ���� '" & pinName & "'."
                        Else
                            e3App.PutInfo 0, "������ " & currentRow & ": ������� '��� - ������� ���������' ������� ���������� � '" & plcUnit & "'."
                        End If
                    Else
                        e3App.PutInfo 0, "������ " & currentRow & ": �������� ��� '��� - ������� ���������' � Excel �����, ������� �� �������."
                    End If

                    On Error GoTo 0 ' ��������� ��������� ������ ����� ��������� ���������
                End If ' End If foundPinId = 0
            End If ' End If foundDeviceId = 0
        End If ' End If deviceName = "" Or pinName = ""
    Next ' Next currentRow

    ' --- ������������ ������� Excel �������� ---
    Call CleanUpExcelObjects(objSheet, objWorkbook, objExcel)
    
    ' --- ������� ��������� � ���������� ������� ����� �������� �������� E3.series ---
    e3App.PutInfo 0, "������ ��������."

    ' --- ������� �������� E3.series ---
    Call CleanUpE3Objects(pin, device, job, e3App)
    
End Sub ' End Sub Main()

' --- ��������������� ������������ ��� ������� �������� ---
Sub CleanUpExcelObjects(ByRef objSheet, ByRef objWorkbook, ByRef objExcel)
    On Error Resume Next ' �������� ��������� ������ ��� �������
    If Not objWorkbook Is Nothing Then objWorkbook.Close False ' ��������� ����� ��� ����������
    If Not objExcel Is Nothing Then objExcel.Quit ' ��������� ���������� Excel
    Set objSheet = Nothing
    Set objWorkbook = Nothing
    Set objExcel = Nothing
    On Error GoTo 0 ' ��������� ��������� ������
End Sub

Sub CleanUpE3Objects(ByRef pin, ByRef device, ByRef job, ByRef e3App)
    On Error Resume Next ' �������� ��������� ������ ��� �������
    Set pin = Nothing
    Set device = Nothing
    Set job = Nothing
    Set e3App = Nothing
    On Error GoTo 0 ' ��������� ��������� ������
End Sub