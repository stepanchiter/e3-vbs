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

' ������� ��� �������� ������������ ���������� � ��������� E_Inom
' ����: ��� ���������� (String)
' ��������: ������ Double(2) - [�������������������, ��������������������]
Dim componentMap

' --- �������� ��������� ������� ---
Sub Main()
    ' ������������� �������� E3.series
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()

    ' ������������� ����������� ������� ��� �������� ��������� OOS ��������
    Set global_foundOOSIds = CreateObject("Scripting.Dictionary")
    
    ' ������������� � ���������� ������� ������������ �����������
    Set componentMap = CreateObject("Scripting.Dictionary")
    Call PopulateComponentMap() ' �������� ��������� ��� ���������� ������� ������������

    e3App.PutInfo 0, "=== ����� �������: ����� OOS �������� � ��������� � ���� ��������� ==="

    ' ��� 1: ������� � ��������� OOS ������� �� �������� ���������
    Call FindAndLogOOSSymbols()

    ' ��� 2: ������� � ������� ���������� � ��������� ����������� (-QF � -KM)
    Call FindAndLogRelatedDevices()

    ' ��� 3: ��������� ��������� QF �� ������ �������� OOS �������
    Call UpdateQFComponentsBasedOnOOSAttribute()

    e3App.PutInfo 0, "=== ���������� ������� ==="

    ' ������� ���������� �������� ��� ������������ ��������
    Call CleanUpGlobalObjects()
End Sub

' --- ��������� ��� ���������� ������� ������������ ����������� ---
Sub PopulateComponentMap()
    ' ��� ������� �������� �������: Add "�������������", Array(�������������������, ��������������������)
    ' ���������� CDbl() ��� ������ �������������� ����� � ��������� ������.
    ' ������������ �������� � ��������� "�� X �� Y" ����� ������������ Y-0.0001 ��� ���������� ������ If-ElseIf
    
    ' �������_3P_0.16-0.25 - �� 0,16 �� 0,2499;  
    componentMap.Add "�������_3P_0.16-0.25A", Array(CDbl(0.16), CDbl(0.2499))
    ' �������_3P_0.25-0.4 - �� 0,25 �� 0,3999;  
    componentMap.Add "�������_3P_0.25-0.4A", Array(CDbl(0.25), CDbl(0.3999))
    ' �������_3P_0.4-0.63 - �� 0,40 �� 0,6299;  
    componentMap.Add "�������_3P_0.4-0.63A", Array(CDbl(0.40), CDbl(0.6299))
    ' �������_3P_0.63-1.0 - �� 0,63 �� 0,9999;  
    componentMap.Add "�������_3P_0.63-1.0A", Array(CDbl(0.63), CDbl(0.9999))
    ' �������_3P_1.0-1.6 - �� 1,00 �� 1,5999;  
    componentMap.Add "�������_3P_1.0-1.6A", Array(CDbl(1.00), CDbl(1.5999))
    ' �������_3P_1.6-2.5 - �� 1,60 �� 2,4999;  
    componentMap.Add "�������_3P_1.6-2.5A", Array(CDbl(1.60), CDbl(2.4999))
    ' �������_3P_2.5-4.0 - �� 2,50 �� 3,9999;  
    componentMap.Add "�������_3P_2.5-4.0A", Array(CDbl(2.50), CDbl(3.9999))
    ' �������_3P_4.0-6.3 - �� 4,00 �� 6,2999;  
    componentMap.Add "�������_3P_4.0-6.3A", Array(CDbl(4.00), CDbl(6.2999))
    ' �������_3P_6.3-10.0 - �� 6,30 �� 9,9999;  
    componentMap.Add "�������_3P_6.3-10.0A", Array(CDbl(6.30), CDbl(9.9999))
    ' �������_3P_9-14 - �� 9,00 �� 13,9999;  
    componentMap.Add "�������_3P_9-14A", Array(CDbl(9.00), CDbl(13.9999))
    ' �������_3P_13-18 - �� 13,00 �� 17,9999;  
    componentMap.Add "�������_3P_13-18A", Array(CDbl(13.00), CDbl(17.9999))
    ' �������_3P_17-23 - �� 17,00 �� 22,9999;  
    componentMap.Add "�������_3P_17-23A", Array(CDbl(17.00), CDbl(22.9999))
    ' �������_3P_20-25 - �� 20,00 �� 24,9999;  
    componentMap.Add "�������_3P_20-25A", Array(CDbl(20.00), CDbl(24.9999))
    ' �������_3P_24-32 - �� 24,00 �� 31,9999;  
    componentMap.Add "�������_3P_24-32A", Array(CDbl(24.00), CDbl(31.9999))
    ' �������_3P_25-40 - �� 25,00 �� 39,9999;  
    componentMap.Add "�������_3P_25-40A", Array(CDbl(25.00), CDbl(39.9999))
    ' �������_3P_40-63 - �� 40,00 �� 62,9999;  
    componentMap.Add "�������_3�_40-63�", Array(CDbl(40.00), CDbl(62.9999))
    ' �������_3P_56-80 - �� 56,00 �� 79,9999; 
    componentMap.Add "�������_3P_56-80A", Array(CDbl(56.00), CDbl(79.9999))
    
    e3App.PutInfo 0, "��������� " & componentMap.Count & " ������������ �����������."
End Sub

' --- ��������� ��� ������ � ������ ���������� �� OOS �������� ---
Sub FindAndLogOOSSymbols()
    Dim symbol          ' ������ Symbol ��� ������ � ���������� ���������
    Dim allSymbolIds()  ' ������ ��� �������� ��������������� ���� �������� � �������
    Dim allSymbolCount  ' ����� ���������� �������� � �������
    Dim i               ' ������� ����� ��� �������� ��������

    Dim symbolName      ' ��� �������� �������
    Dim dProizv3Value   ' �������� �������� "�� D_Proizv3" �������� �������
    Dim OOSIndex        ' �������� ������ �� ����� OOS ������� (��������, 123 ��� "OOS123")

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
    
    e3App.PutInfo 0, "������� " & allSymbolCount & " �������� � �������. ���� OOS ������� � '�� D_Proizv3' = '3' ��� '4'..."

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

            ' ���������, ������������� �� �������� �������� ����� ��������� ("3" ��� "4")
            If dProizv3Value = "3" Or dProizv3Value = "4" Or dProizv3Value = "8" Then
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
        e3App.PutInfo 0, "�� ������� OOS �������� �� ��������� �������� '�� D_Proizv3' ������ '3' ��� '4'."
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
        e3App.PutInfo 0, "��� ��������������� OOS �������� (D_Proizv3=3 ��� 4) ��� ������ ��������� ���������."
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
                    e3App.PutInfo 0, "    ������o -QF ����������: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), �� ��� ��������� ('" & componentName & "') �� �������� '�������'. ��� ���������� ���������."
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
                kmFoundCount = kmFoundCount + 1
                e3App.PutInfo 0, "    ������� -KM ����������: '" & currentDeviceName & "'" & _
                                 " (ID: " & allDeviceIds(i) & ")" & _
                                 " | ���������: '" & componentName & "'"
                foundRelatedDeviceForCurrentOOS = True
            End If
        Next
        
        If kmFoundCount = 0 Then
            e3App.PutInfo 0, "    -KM" & OOSIndex_str & " �� ������� �� ������ ���������� ����� ���� ��������� �������."
        Else
            e3App.PutInfo 0, "    ����� ������� " & kmFoundCount & " -KM ��������� ��� OOS" & OOSIndex_str & "."
        End If

        If Not foundRelatedDeviceForCurrentOOS Then
            e3App.PutInfo 0, "  ��� OOS" & OOSIndex_str & " �� ������� �� ������ ���������������� -QF (� ����������� '�������') ��� -KM ����������."
        End If
    Next

    e3App.PutInfo 0, "=== ���������� ������ ��������� ��������� ==="

    Set device = Nothing ' ����������� ������ Device
End Sub

' --- ��������� ��� ���������� ���������� QF �� ������ �������� OOS ������� ---
Sub UpdateQFComponentsBasedOnOOSAttribute()
    Dim symbolObj       ' ������ Symbol ��� ������ ��������� OOS
    Dim deviceObj       ' ������ Device ��� ���������� ����������� QF
    Dim OOSIndex_str    ' ��������� ������������� ��������� ������� OOS �������
    Dim OOSSymbolId     ' ID OOS �������
    Dim eInomValue_str  ' ��������� �������� �������� "�� E_Inom" (��������)
    Dim eInomValue_num  ' �������� �������� �������� "�� E_Inom"
    Dim isEInomValueValid ' ���� ��� �������� ���������� ��������������
    
    Dim targetDeviceName_QF ' ��������� ��� QF ����������
    Dim allDeviceIds()      ' ������ ID ���� ���������
    Dim allDeviceCount      ' ���������� ���� ���������
    Dim i                   ' ������� �����

    ' ����� ���������� ��� ������ ����������
    Dim componentName_to_set ' ��� ����������, ������� ����� ����������
    Dim rangeValues          ' ������ � ���/���� ���������� ��� �������� ���������� �� �������
    Dim foundMatchingComponent ' ����, �����������, ������ �� ���������� ���������
    Dim componentName_key    ' ���������� ��� �������� ������ �������

    ' ��������� ��� ������ ����������
    Const COMPONENT_VERSION = "1" ' ������ ����������

    ' ���������, ���� �� ������� OOS ������� �� ���������� ����
    If global_foundOOSIds.Count = 0 Then
        e3App.PutInfo 0, "COMM: ��� ��������������� OOS �������� (D_Proizv3=3 ��� 4) ��� ���������� ����������� QF."
        Exit Sub
    End If

    Set symbolObj = job.CreateSymbolObject()
    Set deviceObj = job.CreateDeviceObject()

    e3App.PutInfo 0, "=== ������ ���������� ����������� -QF �� ������ �������� '�� E_Inom' OOS �������� ==="

    ' ���������, ��� ������� ������������ ����������� �� ����
    If componentMap.Count = 0 Then
        e3App.PutInfo 0, "������: ������� ������������ ����������� ����. ���������� �������� ����������."
        Exit Sub
    End If

    ' �������� ������ ���� ��������� � ������� ���� ��� ��� �������������
    allDeviceCount = job.GetAllDeviceIds(allDeviceIds)

    ' ���������� ������ ��������������� OOS ������
    For Each OOSIndex_str In global_foundOOSIds.Keys
        OOSSymbolId = global_foundOOSIds.Item(OOSIndex_str)
        
        ' ������������� OOS ������ ��� ������ ��������
        symbolObj.SetId(OOSSymbolId)
        
        ' ��������� ������� ������ ��� ��������� � �������������� �������� E_Inom
        eInomValue_str = CStr(symbolObj.GetAttributeValue("�� E_Inom")) 

        e3App.PutInfo 0, "  ��������� OOS" & OOSIndex_str & " (ID: " & OOSSymbolId & ")"
        e3App.PutInfo 0, "    �������� ������� '�� E_Inom': '" & eInomValue_str & "'"
        e3App.PutInfo 0, "    ������� '�� E_Inom' ����� Trim(): '" & Trim(eInomValue_str) & "'"

        isEInomValueValid = False ' ���������� ������� ����������
        
        Dim trimmedEInomValue_str : trimmedEInomValue_str = Trim(eInomValue_str)

        If IsNumeric(trimmedEInomValue_str) And Len(trimmedEInomValue_str) > 0 Then
            On Error Resume Next ' �������� ��������� ������ ��� CDbl
            eInomValue_num = CDbl(trimmedEInomValue_str)
            If Err.Number = 0 Then
                isEInomValueValid = True ' �������������� �������
                e3App.PutInfo 0, "    �������: ��������������� �������� �������� '�� E_Inom': " & eInomValue_num
            Else
                e3App.PutInfo 0, "    ������: CDbl �� ������� ������������� ������ '" & trimmedEInomValue_str & "' � ����� (Err: " & Err.Description & ")"
                Err.Clear ' ������� ������
            End If
            On Error GoTo 0 ' ��������� ��������� ������
        Else
            e3App.PutInfo 0, "    ��������: ������� '�� E_Inom' ('" & eInomValue_str & "') ���� ��� �� �������� ������. ���������� ����������."
        End If

        ' ������ ���� �������������� ������ �������, ���� ���������� ���������
        If isEInomValueValid Then
            foundMatchingComponent = False
            componentName_to_set = "" ' ���������� ��� ������� OOS �������

            ' ���������� ������� � ������� ����������� ���������
            For Each componentName_key In componentMap.Keys
                rangeValues = componentMap.Item(componentName_key) ' �������� ������ [min, max]
                
                If eInomValue_num >= rangeValues(0) And eInomValue_num <= rangeValues(1) Then
                    componentName_to_set = componentName_key ' ����� ���������� ��� ����������
                    foundMatchingComponent = True
                    e3App.PutInfo 0, "    ������� ���������� ��� ����������: '" & componentName_to_set & "' ��� �������� " & eInomValue_num
                    Exit For ' ������� �� �����, ��� ��� ����� ������ ����������
                End If
            Next

            If foundMatchingComponent Then
                e3App.PutInfo 0, "    ����� ��������� -QF ��������� ��� ���������� ���������� ��: '" & componentName_to_set & "'..."
                
                targetDeviceName_QF = "-QF" & OOSIndex_str
                Dim qfUpdatedCount : qfUpdatedCount = 0 

                For i = 1 To allDeviceCount
                    deviceObj.SetId(allDeviceIds(i))
                    Dim currentDeviceName : currentDeviceName = deviceObj.GetName()
                    Dim currentComponentName : currentComponentName = deviceObj.GetComponentName()

                    If UCase(currentDeviceName) = UCase(targetDeviceName_QF) Then
                        ' �������������� ��������, ��� ��������� QF �������� "�������"
                        If InStr(1, LCase(currentComponentName), "�������") > 0 Then
                            e3App.PutInfo 0, "      ������� -QF ���������� ��� ����������: '" & currentDeviceName & "'" & _
                                             " (ID: " & allDeviceIds(i) & ", ������� ���������: '" & currentComponentName & "')"
                            
                            On Error Resume Next ' �������� ��������� ������ ��� SetComponentName
                            deviceObj.SetComponentName componentName_to_set, COMPONENT_VERSION
                            If Err.Number = 0 Then
                                qfUpdatedCount = qfUpdatedCount + 1
                                e3App.PutInfo 0, "        �������: ��������� �������� ��: '" & componentName_to_set & "' (������: '" & COMPONENT_VERSION & "')."
                            Else
                                e3App.PutInfo 0, "        ������ ��� ���������� ���������� ��� '" & currentDeviceName & "': " & Err.Description
                                Err.Clear ' ������� ������
                            End If
                            On Error GoTo 0 ' ��������� ��������� ������
                        Else
                            e3App.PutInfo 0, "      ������o -QF ����������: '" & currentDeviceName & "' (ID: " & allDeviceIds(i) & "), �� ��� ��������� ('" & currentComponentName & "') �� �������� '�������'. ��������� ����������."
                        End If
                    End If
                Next
                
                If qfUpdatedCount = 0 Then
                    e3App.PutInfo 0, "    ��� OOS" & OOSIndex_str & " �� ������� �� ������ -QF ���������� � ����������� '�������' ��� ����������."
                Else
                    e3App.PutInfo 0, "    ����� ��������� " & qfUpdatedCount & " -QF ��������� ��� OOS" & OOSIndex_str & "."
                End If
            Else
                e3App.PutInfo 0, "    ��������: ��� �������� '�� E_Inom' (" & eInomValue_num & ") �� ������� ����������� ���������� � ������� ������������. ���������� ���������."
            End If
        End If 
    Next ' ���������� � ���������� OOS �������

    e3App.PutInfo 0, "=== ���������� ���������� ����������� -QF ==="

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
    ' ����������� ������ componentMap
    If Not componentMap Is Nothing Then
        Set componentMap = Nothing
    End If
End Sub

' --- ����� ����� � ������: ��������� �������� ��������� ---
Call Main()