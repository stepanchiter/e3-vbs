'*******************************************************************************
' �������� �������: E3_ComponentUpdater_Combined
' �����: E3.series VBScript Assistant
' ����: 08.07.2025
' ��������: ������ ��� ��������������� ���������� ���� ����������� ��� ��������� -QF � -KM
'          �� ������ ���������, ����������� �� OOS ��������, � ����� ������� ������������.
'          ������������� ���:
'          1. ������ OOS �������� �� �������� "�� D_Proizv3" = "2" ��� "7".
'          2. ��������� �������� "�� E_Inom" �� 1.25 ��� "�� D_Proizv3" = "2"
'             � �� 1.35 ��� "�� D_Proizv3" = "7" ����� ��������������.
'          3. ������������� ����� ������� ������������ ����������� ��� -QF.
'          4. ���������� ���������� ����������� ��� ��������� -KM � �������������� ��������� ������� ������������.
'          5. ���������� ������ "��� ���� ��������������" ����� �������� ��������� ���������� Dim.
'          6. ���������� ������ "�������������� ������� ����������" ����� ���������������� ���� ��� GoTo.
'*******************************************************************************
Option Explicit

' --- ���������� ���������� ---
' ������ ���������� E3.series
Dim e3App
' ������ Job, �������������� ������� ������
Dim job
' ������� ��� �������� ID ��������� OOS ��������, ��������������� ���������.
' ����: ��������� �������� �� ����� OOS ������� (��������, "123" ��� "OOS123")
' ��������: ID ������� � E3.series
Dim global_foundOOSIds

' ������� ��� �������� ������������ ���������� QF � ��������� E_Inom
' ����: ��� ���������� (String)
' ��������: ������ Double(2) - [�������������������, ��������������������]
Dim qfComponentMap

' ������� ��� �������� ������������ ���������� KM � ��������� E_Inom
Dim kmComponentMap

' --- �������� ��������� ������� ---
Sub Main()
    ' ������������� �������� E3.series
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()

    ' ������������� ����������� ������� ��� �������� ��������� OOS ��������
    Set global_foundOOSIds = CreateObject("Scripting.Dictionary")
    
    ' ������������� � ���������� �������� ������������ �����������
    Set qfComponentMap = CreateObject("Scripting.Dictionary")
    Call PopulateQFComponentMap() ' ��������� ��� ���������� ������� ������������ QF

    Set kmComponentMap = CreateObject("Scripting.Dictionary") ' ������������� ������ �������
    Call PopulateKMComponentMap() ' ��������� ��� ���������� ������� ������������ KM

    e3App.PutInfo 0, "=== ����� �������: ����� OOS �������� � ��������� � ���� ��������� ==="

    ' ��� 1: ������� � ��������� OOS ������� �� �������� ���������
    Call FindAndLogOOSSymbols()

    ' ��� 2: ������� � ������� ���������� � ��������� ����������� (-QF � -KM)
    ' ����������: ����������� ��� -KM ����� ����� ���������� ��������� "���������"
    Call FindAndLogRelatedDevices()

    ' ��� 3: ��������� ���������� QF � KM �� ������ �������� OOS �������
    ' ����������: �������� ��������� E_Inom � ����������� �� D_Proizv3
    Call UpdateComponentsBasedOnOOSAttribute()

    e3App.PutInfo 0, "=== ���������� ������� ==="

    ' ������� ���������� �������� ��� ������������ ��������
    Call CleanUpGlobalObjects()
End Sub

' --- ��������� ��� ���������� ������� ������������ ����������� QF ---
Sub PopulateQFComponentMap()
    qfComponentMap.Add "�������_3�_10A_13176DEK", Array(CDbl(0.01), CDbl(10.00))
    qfComponentMap.Add "�������_3�_16A_13177DEK", Array(CDbl(10.01), CDbl(16.00))
    qfComponentMap.Add "�������_3�_20A_13178DEK", Array(CDbl(16.01), CDbl(20.00))
    qfComponentMap.Add "�������_3�_25A_13179DEK", Array(CDbl(20.01), CDbl(25.00))
    qfComponentMap.Add "�������_3�_32A_13180DEK", Array(CDbl(25.01), CDbl(32.00))
    qfComponentMap.Add "�������_3�_40A_13181DEK", Array(CDbl(32.01), CDbl(40.00))
    qfComponentMap.Add "�������_3�_50A_13182DEK", Array(CDbl(40.01), CDbl(50.00))
    qfComponentMap.Add "�������_3�_63A_13183DEK", Array(CDbl(50.01), CDbl(63.00))
    qfComponentMap.Add "�������_3�_80A_13008DEK", Array(CDbl(63.01), CDbl(80.00))
    qfComponentMap.Add "�������_3�_100A_13009DEK", Array(CDbl(80.01), CDbl(100.00))
    qfComponentMap.Add "�������_3�_125A_13027DEK", Array(CDbl(100.01), CDbl(125.00))
    qfComponentMap.Add "�������_3�_160A_22752DEK", Array(CDbl(125.01), CDbl(160.00))
    qfComponentMap.Add "�������_3�_200A_22754DEK", Array(CDbl(160.01), CDbl(200.00))
    qfComponentMap.Add "�������_3�_250A_22756DEK", Array(CDbl(200.01), CDbl(250.00))
    
    e3App.PutInfo 0, "��������� " & qfComponentMap.Count & " ������������ ����������� ��� -QF."
End Sub

' --- ��������� ��� ���������� ������� ������������ ����������� KM ---
Sub PopulateKMComponentMap()
    kmComponentMap.Add "���������_��102_22001DEK", Array(CDbl(0.00), CDbl(9.00))
    kmComponentMap.Add "���������_��102_22002DEK", Array(CDbl(9.01), CDbl(12.00))
    kmComponentMap.Add "���������_��102_22003DEK", Array(CDbl(12.01), CDbl(18.00))
    kmComponentMap.Add "���������_��102_22004DEK", Array(CDbl(18.01), CDbl(25.00))
    kmComponentMap.Add "���������_��102_22005DEK", Array(CDbl(25.01), CDbl(32.00))
    kmComponentMap.Add "���������_��102_22006DEK", Array(CDbl(32.01), CDbl(40.00))
    kmComponentMap.Add "���������_��102_22007DEK", Array(CDbl(40.01), CDbl(50.00))
    kmComponentMap.Add "���������_��102_22008DEK", Array(CDbl(50.01), CDbl(65.00))
    kmComponentMap.Add "���������_��102_22009DEK", Array(CDbl(65.01), CDbl(80.00))
    kmComponentMap.Add "���������_��102_22010DEK", Array(CDbl(80.01), CDbl(95.00))
    kmComponentMap.Add "���������_��103_22150DEK", Array(CDbl(95.01), CDbl(115.00))
    kmComponentMap.Add "���������_��103_22152DEK", Array(CDbl(115.01), CDbl(150.00))
    kmComponentMap.Add "���������_��103_22154DEK", Array(CDbl(150.01), CDbl(185.00))
    kmComponentMap.Add "���������_��103_22156DEK", Array(CDbl(185.01), CDbl(225.00))
    kmComponentMap.Add "���������_��103_22158DEK", Array(CDbl(225.01), CDbl(265.00))
    kmComponentMap.Add "���������_��103_22160DEK", Array(CDbl(265.01), CDbl(330.00))
    kmComponentMap.Add "���������_��103_22162DEK", Array(CDbl(330.01), CDbl(400.00))

    e3App.PutInfo 0, "��������� " & kmComponentMap.Count & " ������������ ����������� ��� -KM."
End Sub

' --- ��������� ��� ������ � ������ ���������� �� OOS �������� ---
Sub FindAndLogOOSSymbols()
    Dim symbol            ' ������ Symbol ��� ������ � ���������� ���������
    Dim allSymbolIds()    ' ������ ��� �������� ��������������� ���� �������� � �������
    Dim allSymbolCount    ' ����� ���������� �������� � �������
    Dim i                 ' ������� ����� ��� �������� ��������

    Dim symbolName        ' ��� �������� �������
    Dim dProizv3Value     ' �������� �������� "�� D_Proizv3" �������� �������
    Dim OOSIndex          ' �������� ������ �� ����� OOS ������� (��������, 123 ��� "OOS123")

    ' ������� ������ Symbol
    Set symbol = job.CreateSymbolObject()

    ' �������� ������ ���� �������� � ������� �������
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    ' ���������, ���� �� ������� � �������
    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "� ������� ������� �� ������� �������� ��� �������."
        Set symbol = Nothing ' ����������� ������ Symbol ����� �������
        Exit Sub
    End If    
    
    e3App.PutInfo 0, "������� " & allSymbolCount & " �������� � �������. ���� OOS ������� � '�� D_Proizv3' = '2' ��� '7'..."

    Dim foundOOSCount : foundOOSCount = 0 ' ������� ��������� OOS ��������, ��������������� ���������

    ' ���������� ��� ������� � �������
    For i = 1 To allSymbolCount
        ' ������������� ������� ������ �� ��� ID ��� ���������� ������
        symbol.SetId(allSymbolIds(i))
        symbolName = symbol.GetName() ' �������� ��� �������

        ' ���������, ���������� �� ��� ������� � "OOS" (��� ����� ��������)
        If LCase(Left(symbolName, 3)) = "OOS" Then
            ' �������� �������� �������� "�� D_Proizv3"
            ' Trim() ������� ������ �������, CStr() ����������� � ������ ��� ��������� ���������
            dProizv3Value = Trim(CStr(symbol.GetAttributeValue("�� D_Proizv3")))

            ' ����������: ���������, ������������� �� �������� �������� ����� ��������� ("2" ��� "7")
            If dProizv3Value = "2" Or dProizv3Value = "7" Then
                foundOOSCount = foundOOSCount + 1 ' ����������� �������
                
                ' ��������� �������� ������ �� ����� OOS ������� (��������, "123" �� "OOS123")
                On Error Resume Next ' �������� ��������� ������ ��� CLng
                OOSIndex = CLng(Mid(symbolName, 4)) ' �������� ������������� ����� ����� � �����
                If Err.Number <> 0 Then
                    ' ���� �������������� �� ������� (��������, "OOSABC"), ���������� �������� ������
                    OOSIndex = Mid(symbolName, 4)
                    e3App.PutInfo 0, "    ��������: �� ������� ������������� ������ '" & Mid(symbolName, 4) & "' � ����� ��� OOS ������� '" & symbolName & "'."
                    Err.Clear ' ������� ������
                End If
                On Error GoTo 0 ' ��������� ��������� ������

                ' ��������� ��������� ������ � ���������� �������
                ' ���������� CStr(OOSIndex) ��� �����, ����� ���� ���������� � ���� ������ �����
                If Not global_foundOOSIds.Exists(CStr(OOSIndex)) Then
                    global_foundOOSIds.Add CStr(OOSIndex), allSymbolIds(i)
                    e3App.PutInfo 0, "  ������ � �������� OOS ������: '" & symbolName & "'" & _
                                     " (ID: " & allSymbolIds(i) & ")" & _
                                     " | ������� '�� D_Proizv3': '" & dProizv3Value & "'"
                Else
                    e3App.PutInfo 0, "  ��������: OOS ������ � �������� '" & CStr(OOSIndex) & "' ��� ������. ��������� ID ��: " & allSymbolIds(i) & _
                                     " (���: '" & symbolName & "', D_Proizv3: '" & dProizv3Value & "')"
                    global_foundOOSIds.Item(CStr(OOSIndex)) = allSymbolIds(i) ' ��������� ID, ���� ����� ������ ��� ����
                End If
            End If
        End If
    Next

    ' ������� �������� ��������� � ����������� ������ OOS ��������
    If foundOOSCount = 0 Then
        e3App.PutInfo 0, "�� ������� OOS �������� �� ��������� �������� '�� D_Proizv3' ������ '2' ��� '7'."
    Else
        e3App.PutInfo 0, "����� ������� " & foundOOSCount & " OOS ��������, ��������������� �������� ���������."
        e3App.PutInfo 0, "ID ��������� OOS �������� ��������� � ���������� ������� 'global_foundOOSIds'."
    End If

    Set symbol = Nothing ' ����������� ������ Symbol
End Sub

' --- ��������� ��� ������ � ������ ���������� � ��������� ����������� (-QF � -KM) ---
Sub FindAndLogRelatedDevices()
    Dim device          ' ������ Device ��� ������ � ������������
    Dim OOSIndex_str    ' ��������� ������������� ��������� ������� OOS �������
    Dim targetDeviceName    ' ��� ����������, ������� �� ���� (��������, "-QF123" ��� "-KM123")
    Dim allDeviceIds()      ' ������ ��� �������� ��������������� ���� ��������� � �������
    Dim allDeviceCount      ' ����� ���������� ��������� � �������
    Dim i                   ' ������� ����� ��� �������� ���������
    Dim currentDeviceName   ' ��� �������� ����������
    Dim componentName       ' ��� ���������� �������� ����������

    ' ���������, ���� �� ������� OOS ������� �� ���������� ����
    If global_foundOOSIds.Count = 0 Then
        e3App.PutInfo 0, "��� ��������������� OOS �������� (D_Proizv3=2 ��� 7) ��� ������ ��������� ���������."
        Exit Sub
    End If

    Set device = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== ������ ������ ��������� ��������� -QF � -KM ��� ��������� OOS �������� ==="
    
    ' �������� ������ ���� ��������� � ������� ���� ��� ��� �������������
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' ���������� ������ ��������������� OOS ������
    For Each OOSIndex_str In global_foundOOSIds.Keys
        e3App.PutInfo 0, "  ����� ��������� ��������� ��� OOS" & OOSIndex_str & ":"
        Dim foundRelatedDeviceForCurrentOOS : foundRelatedDeviceForCurrentOOS = False

        ' --- ����� -QF ��������� ---
        targetDeviceName = "-QF" & OOSIndex_str
        Dim qfFoundCount : qfFoundCount = 0 ' ������� ��������� �������� -QF ���������
        
        For i = 1 To allDeviceCount ' ���������� ��� ����������, ����� ����� ��� ����������
            device.SetId(allDeviceIds(i))
            currentDeviceName = device.GetName()
            componentName = device.GetComponentName()

            If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                ' ��� -QF ���������, ����� ���������, �������� �� ��������� "�������"
                If InStr(1, LCase(componentName), "�������") > 0 Then
                    qfFoundCount = qfFoundCount + 1
                    e3App.PutInfo 0, "    ������� -QF ����������: '" & currentDeviceName & "'" & _
                                     " (ID: " & allDeviceIds(i) & ")" & _
                                     " | ���������: '" & componentName & "'"
                    foundRelatedDeviceForCurrentOOS = True
                Else
                    e3App.PutInfo 0, "    ������� -QF ����������: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), �� ��� ��������� ('" & componentName & "') �� �������� '�������'. ��� ���������� ���������."
                End If
            End If
        Next
        
        If qfFoundCount = 0 Then
            e3App.PutInfo 0, "    -QF" & OOSIndex_str & " (� ����������� '�������') �� ������� �� ������ ���������� ����� ���� ��������� �������."
        Else
            e3App.PutInfo 0, "    ����� ������� " & qfFoundCount & " -QF ��������� � ����������� '�������' ��� OOS" & OOSIndex_str & "."
        End If

        ' --- ����� -KM ��������� ---
        targetDeviceName = "-KM" & OOSIndex_str
        Dim kmFoundCount : kmFoundCount = 0 ' ������� ��������� -KM ���������

        For i = 1 To allDeviceCount ' ���������� ��� ����������, ����� ����� ��� ����������
            device.SetId(allDeviceIds(i))
            currentDeviceName = device.GetName()
            componentName = device.GetComponentName()

            If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                ' ��� -KM ���������, ����� ���������, �������� �� ��������� "���������"
                If InStr(1, LCase(componentName), "���������") > 0 Then
                    kmFoundCount = kmFoundCount + 1
                    e3App.PutInfo 0, "    ������� -KM ����������: '" & currentDeviceName & "'" & _
                                 " (ID: " & allDeviceIds(i) & ")" & _
                                 " | ���������: '" & componentName & "'"
                    foundRelatedDeviceForCurrentOOS = True
                Else
                    e3App.PutInfo 0, "    ������� -KM ����������: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), �� ��� ��������� ('" & componentName & "') �� �������� '���������'. ��� ���������� ���������."
                End If
            End If
        Next
        
        If kmFoundCount = 0 Then
            e3App.PutInfo 0, "    -KM" & OOSIndex_str & " (� ����������� '���������') �� ������� �� ������ ���������� ����� ���� ��������� �������."
        Else
            e3App.PutInfo 0, "    ����� ������� " & kmFoundCount & " -KM ��������� � ����������� '���������' ��� OOS" & OOSIndex_str & "."
        End If

        If Not foundRelatedDeviceForCurrentOOS Then
            e3App.PutInfo 0, "  ��� OOS" & OOSIndex_str & " �� ������� �� ������ ���������������� -QF (� ����������� '�������') ��� -KM (� ����������� '���������') ����������."
        End If
    Next

    e3App.PutInfo 0, "=== ���������� ������ ��������� ��������� ==="

    Set device = Nothing ' ����������� ������ Device
End Sub

' --- ��������� ��� ���������� ����������� QF � KM �� ������ �������� OOS ������� ---
Sub UpdateComponentsBasedOnOOSAttribute()
    Dim symbolObj       ' ������ Symbol ��� ������ ��������� OOS
    Dim deviceObj       ' ������ Device ��� ���������� ����������� QF/KM
    Dim OOSIndex_str    ' ��������� ������������� ��������� ������� OOS �������
    Dim OOSSymbolId     ' ID OOS �������
    Dim eInomValue_str  ' ��������� �������� �������� "�� E_Inom" (��������)
    Dim eInomValue_num  ' �������� �������� �������� "�� E_Inom"
    Dim modifiedEInomValue_num ' �������� �������� "�� E_Inom" ����� ���������
    Dim isEInomValueValid ' ���� ��� �������� ���������� ��������������
    
    Dim targetDeviceName    ' ��������� ��� ���������� (������������ ��� QF � KM)
    Dim allDeviceIds()      ' ������ ID ���� ���������
    Dim allDeviceCount      ' ���������� ���� ���������
    Dim i                   ' ������� �����

    Dim componentName_to_set ' ��� ����������, ������� ����� ����������
    Dim rangeValues          ' ������ � ���/���� ���������� ��� �������� ���������� �� �������
    Dim foundMatchingComponent ' ����, �����������, ������ �� ���������� ���������
    Dim componentName_key    ' ���������� ��� �������� ������ �������
    Dim trimmedEInomValue_str ' ��� ��������� ������ E_Inom

    ' ���������� ��� ��������� ����������� ��������� (����������� ���� ���)
    Dim qfUpdatedCount
    Dim kmUpdatedCount
    
    ' ���������� ��� ���� ������� ��������� (����������� ���� ���)
    Dim currentDeviceName
    Dim currentComponentName

    ' ����� ���������� ��� �������� D_Proizv3 � ������������ ���������
    Dim dProizv3Value     ' �������� �������� "�� D_Proizv3" �������� OOS �������
    Dim multiplicationFactor ' ����������� ��������� ��� E_Inom

    ' ��������� ��� ������ ����������
    Const COMPONENT_VERSION = "1" ' ������ ����������

    ' ���������, ���� �� ������� OOS ������� �� ���������� ����
    If global_foundOOSIds.Count = 0 Then
        e3App.PutInfo 0, "COMM: ��� ��������������� OOS �������� (D_Proizv3=2 ��� 7) ��� ���������� �����������."
        Exit Sub
    End If

    Set symbolObj = job.CreateSymbolObject()
    Set deviceObj = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== ������ ���������� ����������� -QF � -KM �� ������ �������� '�� E_Inom' OOS �������� ==="

    ' ���������, ��� ������� ������������ ����������� �� �����
    If qfComponentMap.Count = 0 And kmComponentMap.Count = 0 Then
        e3App.PutInfo 0, "������: ������� ������������ ����������� QF � KM �����. ���������� �������� ����������."
        Exit Sub
    End If

    ' �������� ������ ���� ��������� � ������� ���� ��� ��� �������������
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' ���������� ������ ��������������� OOS ������
    For Each OOSIndex_str In global_foundOOSIds.Keys
        OOSSymbolId = global_foundOOSIds.Item(OOSIndex_str)
        
        ' ������������� OOS ������ ��� ������ ��������
        symbolObj.SetId(OOSSymbolId)
        
        ' �������� �������� �������� "�� E_Inom" � "�� D_Proizv3"
        eInomValue_str = CStr(symbolObj.GetAttributeValue("�� E_Inom"))    
        dProizv3Value = Trim(CStr(symbolObj.GetAttributeValue("�� D_Proizv3")))

        e3App.PutInfo 0, "  ��������� OOS" & OOSIndex_str & " (ID: " & OOSSymbolId & ")"
        e3App.PutInfo 0, "    �������� ������� '�� E_Inom': '" & eInomValue_str & "'"
        e3App.PutInfo 0, "    ������� '�� D_Proizv3': '" & dProizv3Value & "'"

        isEInomValueValid = False ' ���������� ������� ����������
        multiplicationFactor = 0 ' �������������� �����������

        ' ���������� ����������� ��������� �� ������ D_Proizv3
        If dProizv3Value = "2" Then
            multiplicationFactor = 1.25
            e3App.PutInfo 0, "    ������ ����������� ���������: " & multiplicationFactor & " (��� D_Proizv3 = '2')"
        ElseIf dProizv3Value = "7" Then
            multiplicationFactor = 1.35
            e3App.PutInfo 0, "    ������ ����������� ���������: " & multiplicationFactor & " (��� D_Proizv3 = '7')"
        Else
            e3App.PutInfo 0, "    ��������: ����������� ��� ���������������� �������� '�� D_Proizv3': '" & dProizv3Value & "'. ���������� ���������� ��� ����� OOS �������."
            ' ����� ��� GoTo. ���� multiplicationFactor �� ����������, ��������� ��� �� ����������.
        End If
        
        ' ������ ���� multiplicationFactor ��� ���������� (�.�., D_Proizv3 ��� '2' ��� '7'),
        ' ���������� ��������� E_Inom � ���������� �����������.
        If multiplicationFactor > 0 Then
            trimmedEInomValue_str = Trim(eInomValue_str)

            If IsNumeric(trimmedEInomValue_str) And Len(trimmedEInomValue_str) > 0 Then
                On Error Resume Next ' �������� ��������� ������ ��� CDbl
                eInomValue_num = CDbl(trimmedEInomValue_str)
                If Err.Number = 0 Then
                    isEInomValueValid = True ' �������������� �������
                    e3App.PutInfo 0, "    �������: ��������������� �������� �������� '�� E_Inom': " & eInomValue_num
                    
                    ' �������� �� ������������ �����������
                    modifiedEInomValue_num = eInomValue_num * multiplicationFactor
                    e3App.PutInfo 0, "    �������� '�� E_Inom' ����� ��������� �� " & multiplicationFactor & ": " & modifiedEInomValue_num
                Else
                    e3App.PutInfo 0, "    ������: CDbl �� ������� ������������� ������ '" & trimmedEInomValue_str & "' � ����� (Err: " & Err.Description & ")"
                    Err.Clear ' ������� ������
                End If
                On Error GoTo 0 ' ��������� ��������� ������
            Else
                e3App.PutInfo 0, "    ��������: ������� '�� E_Inom' ('" & eInomValue_str & "') ���� ��� �� �������� ������. ���������� ����������."
            End If

            ' ������ ���� �������������� ������ �������, ���� ���������� ��������� � ���������
            If isEInomValueValid Then
                ' --- ���������� QF ����������� ---
                e3App.PutInfo 0, "    ����� � ���������� -QF ���������..."
                foundMatchingComponent = False
                componentName_to_set = ""
                qfUpdatedCount = 0 ' �������������� ������� ��� �������� OOS �������
                
                ' ���� � ������� QF
                For Each componentName_key In qfComponentMap.Keys
                    rangeValues = qfComponentMap.Item(componentName_key)
                    
                    If modifiedEInomValue_num >= rangeValues(0) And modifiedEInomValue_num <= rangeValues(1) Then
                        componentName_to_set = componentName_key
                        foundMatchingComponent = True
                        e3App.PutInfo 0, "      ������� ���������� ��� ���������� QF: '" & componentName_to_set & "'"
                        Exit For
                    End If
                Next

                If foundMatchingComponent Then
                    targetDeviceName = "-QF" & OOSIndex_str
                    
                    For i = 1 To allDeviceCount
                        deviceObj.SetId(allDeviceIds(i))
                        currentDeviceName = deviceObj.GetName()
                        currentComponentName = deviceObj.GetComponentName()

                        If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                            If InStr(1, LCase(currentComponentName), "�������") > 0 Then
                                On Error Resume Next
                                deviceObj.SetComponentName componentName_to_set, COMPONENT_VERSION
                                If Err.Number = 0 Then
                                    qfUpdatedCount = qfUpdatedCount + 1
                                    e3App.PutInfo 0, "        �������: ��������� -QF '" & currentDeviceName & "' �������� ��: '" & componentName_to_set & "'."
                                Else
                                    e3App.PutInfo 0, "        ������ ��� ���������� ���������� QF ��� '" & currentDeviceName & "': " & Err.Description
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            Else
                                e3App.PutInfo 0, "      ������� -QF ����������: '" & currentDeviceName & "', �� ��� ��������� ('" & currentComponentName & "') �� �������� '�������'. ��������� ����������."
                            End If
                        End If
                    Next
                    If qfUpdatedCount = 0 Then
                        e3App.PutInfo 0, "    ��� OOS" & OOSIndex_str & " �� ������� �� ������ -QF ���������� � ����������� '�������' ��� ����������."
                    Else
                        e3App.PutInfo 0, "    ����� ��������� " & qfUpdatedCount & " -QF ��������� ��� OOS" & OOSIndex_str & "."
                    End If
                Else
                    e3App.PutInfo 0, "    ��������: ��� ����������������� �������� " & modifiedEInomValue_num & " �� ������� ����������� ���������� QF � ������� ������������. ���������� QF ���������."
                End If


                ' --- ���������� KM ����������� ---
                e3App.PutInfo 0, "    ����� � ���������� -KM ���������..."
                foundMatchingComponent = False
                componentName_to_set = ""
                kmUpdatedCount = 0 ' �������������� ������� ��� �������� OOS �������

                ' ���� � ������� KM
                For Each componentName_key In kmComponentMap.Keys
                    rangeValues = kmComponentMap.Item(componentName_key)
                    
                    If modifiedEInomValue_num >= rangeValues(0) And modifiedEInomValue_num <= rangeValues(1) Then
                        componentName_to_set = componentName_key
                        foundMatchingComponent = True
                        e3App.PutInfo 0, "      ������� ���������� ��� ���������� KM: '" & componentName_to_set & "'"
                        Exit For
                    End If
                Next

                If foundMatchingComponent Then
                    targetDeviceName = "-KM" & OOSIndex_str
                    
                    For i = 1 To allDeviceCount
                        deviceObj.SetId(allDeviceIds(i))
                        currentDeviceName = deviceObj.GetName()
                        currentComponentName = deviceObj.GetComponentName()

                        If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                            ' �������������� ��������, ��� ��������� KM �������� "���������"
                            If InStr(1, LCase(currentComponentName), "���������") > 0 Then
                                On Error Resume Next
                                deviceObj.SetComponentName componentName_to_set, COMPONENT_VERSION
                                If Err.Number = 0 Then
                                    kmUpdatedCount = kmUpdatedCount + 1
                                    e3App.PutInfo 0, "        �������: ��������� -KM '" & currentDeviceName & "' �������� ��: '" & componentName_to_set & "'."
                                Else
                                    e3App.PutInfo 0, "        ������ ��� ���������� ���������� KM ��� '" & currentDeviceName & "': " & Err.Description
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            Else
                                e3App.PutInfo 0, "      ������� -KM ����������: '" & currentDeviceName & "', �� ��� ��������� ('" & currentComponentName & "') �� �������� '���������'. ��������� ����������."
                            End If
                        End If
                    Next
                    If kmUpdatedCount = 0 Then
                        e3App.PutInfo 0, "    ��� OOS" & OOSIndex_str & " �� ������� �� ������ -KM ���������� � ����������� '���������' ��� ����������."
                    Else
                        e3App.PutInfo 0, "    ����� ��������� " & kmUpdatedCount & " -KM ��������� ��� OOS" & OOSIndex_str & "."
                    End If
                Else
                    e3App.PutInfo 0, "    ��������: ��� ����������������� �������� " & modifiedEInomValue_num & " �� ������� ����������� ���������� KM � ������� ������������. ���������� KM ���������."
                End If
            End If ' End If isEInomValueValid Then
        End If ' End If multiplicationFactor > 0 Then
    Next ' ���������� � ���������� OOS �������

    e3App.PutInfo 0, "=== ���������� ���������� ����������� -QF � -KM ==="

    Set symbolObj = Nothing ' ����������� ������ Symbol
    Set deviceObj = Nothing ' ����������� ������ Device
End Sub


' --- ��������������� ��������� ��� ������� ���������� �������� ---
Sub CleanUpGlobalObjects()
    ' ���������, ��� ������� ����������, ������ ��� �� �����������
    If Not job Is Nothing Then
        Set job = Nothing
    End If
    If Not e3App Is Nothing Then
        Set e3App = Nothing
    End If
    If Not global_foundOOSIds Is Nothing Then
        Set global_foundOOSIds = Nothing
    End If
    ' ����������� ������� �������� �����������
    If Not qfComponentMap Is Nothing Then
        Set qfComponentMap = Nothing
    End If
    If Not kmComponentMap Is Nothing Then
        Set kmComponentMap = Nothing
    End If
End Sub

' --- ����� ����� � ������: ��������� �������� ��������� ---
Call Main()