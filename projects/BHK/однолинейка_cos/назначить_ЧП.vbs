'*******************************************************************************
' �������� �������: E3_UZ_ComponentUpdater
' �����: E3.series VBScript Assistant
' ����: 01.07.2025
' ��������: ������ ��� ��������������� ���������� ���� ����������� ��� ��������� -UZ
'          �� ������ ���������, ����������� �� OOS ��������, � ����� ������� ������������.
'          �� ��������� ��� ���������� �� "�������".
'*******************************************************************************

Option Explicit

'*******************************************************************************
' ���������� ����������
'*******************************************************************************
Dim global_e3App             ' ������ ���������� E3.series
Dim global_job               ' ������ �������� ������� E3.series
Dim global_OOSArticles       ' ������� ��� �������� �������� OOS � ����������� ��������� (OOSIndex -> ExtractedArticle)
Dim global_articleComponentMap ' ������� ��� �������� ������������ ������� -> ����� ��� ���������� UZ

'*******************************************************************************
' ��������� ExtractOOSArticles()
' ���� ������� OOS, ��������� �� ��� �������� � ��������� � global_OOSArticles.
'*******************************************************************************
Sub ExtractOOSArticles()
    Dim symbolIds, symbolCount, i
    Dim symbol, symbolName, attributeValue, extractedArticle, OOSIndex
    Dim regEx, matches

    Set symbol = global_job.CreateSymbolObject()
    Set regEx = New RegExp
    regEx.IgnoreCase = True ' ��� ������ OOS, ABA, ABC ��� ����� ��������
    regEx.Global = True

    global_e3App.PutInfo 0, "�������� ����� OOS �������� � ���������� ���������..."

    ' �����������: ���������� GetSymbolIds ������ GetAllSymbolIds
    symbolCount = global_job.GetSymbolIds(symbolIds) 

    If symbolCount > 0 Then
        For i = 1 To symbolCount
            symbol.SetId(symbolIds(i))
            symbolName = symbol.GetName()

            ' ���������, �������� �� ������ OOS �������� (��������, OOS1, OOS10)
            If LCase(Left(symbolName, 3)) = "OOS" Then
                ' ��������� ������� "�� D_Proizv3" (Manufacturer Specific Data 3)
                attributeValue = symbol.GetAttributeValue("�� D_Proizv3")

                If attributeValue = "1" Then
                    ' ��������� ������� �� �������� "�� D_Proizv2" (Manufacturer Specific Data 2)
                    attributeValue = symbol.GetAttributeValue("�� D_Proizv2")

                    regEx.Pattern = "(ABA|ABC)\d+" ' ������� ��� ���������
                    Set matches = regEx.Execute(attributeValue)

                    If matches.Count > 0 Then
                        extractedArticle = matches(0).Value
                        ' ��������� �������� ������ �� ����� ������� OOS (��������, �� OOS123 -> 123)
                        regEx.Pattern = "\d+"
                        Set matches = regEx.Execute(symbolName)
                        If matches.Count > 0 Then
                            OOSIndex = CInt(matches(0).Value)
                            If Not global_OOSArticles.Exists(OOSIndex) Then
                                global_OOSArticles.Add OOSIndex, extractedArticle
                                global_e3App.PutInfo 0, "������ OOS ������: " & symbolName & _
                                    ", ������� '�� D_Proizv3' = '" & symbol.GetAttributeValue("�� D_Proizv3") & _
                                    "', �������� ������� �� '�� D_Proizv2': '" & extractedArticle & "'"
                            Else
                                global_e3App.PutInfo 1, "��������������: OOS ������ '" & OOSIndex & _
                                    "' ��� ���������� � �������. ���������� ��������: " & symbolName
                            End If
                        Else
                            global_e3App.PutInfo 1, "��������������: �� ������� ������� �������� ������ �� ����� OOS �������: " & symbolName
                        End If
                    Else
                        global_e3App.PutInfo 1, "��������������: � �������� '�� D_Proizv2' ������� '" & symbolName & _
                            "' �� ������ ������� �� ������� (ABA|ABC)######."
                    End If
                Else
                    global_e3App.PutInfo 0, "���������� ������ '" & symbolName & _
                        "', ��� ��� ������� '�� D_Proizv3' �� ����� '1'."
                End If
            End If
        Next
    Else
        global_e3App.PutInfo 1, "� ������� �� ������� ��������."
    End If

    global_e3App.PutInfo 0, "����� OOS �������� ��������. ������� ���������: " & global_OOSArticles.Count
    Call CleanUpObjects(regEx, matches, symbol, Nothing, Nothing) ' matches � symbol ��������� �����
End Sub

'*******************************************************************************
' ��������� CleanUpObjects()
' ����������� COM-�������.
'*******************************************************************************
Sub CleanUpObjects(ByRef reObj, ByRef matchesObj, ByRef symbolObj, ByRef jobObj, ByRef appObj)
    If Not reObj Is Nothing Then Set reObj = Nothing
    If Not matchesObj Is Nothing Then Set matchesObj = Nothing
    If Not symbolObj Is Nothing Then Set symbolObj = Nothing
    If Not jobObj Is Nothing Then Set jobObj = Nothing
    If Not appObj Is Nothing Then Set appObj = Nothing
End Sub

'*******************************************************************************
' ��������� FindUZDevices(job)
' ���� ���������� -UZ### � ������� � ��� ����������.
' (��� ��������� ������ ��� �����������/����������� � �� �������� ������ ��� ���������.)
'*******************************************************************************
Sub FindUZDevices(job)
    Dim deviceIds, deviceCount, i
    Dim device, devName
    Dim regEx, deviceInfo ' deviceInfo - ��������� ������� ��� ���� ���������

    Set device = job.CreateDeviceObject()
    Set regEx = New RegExp
    regEx.Pattern = "^-UZ\d+$" ' ������� ��� ������ -UZ###
    regEx.IgnoreCase = True
    Set deviceInfo = CreateObject("Scripting.Dictionary")

    global_e3App.PutInfo 0, "�������� ����� ��������� -UZ..."

    deviceCount = job.GetAllDeviceIds(deviceIds)

    If deviceCount > 0 Then
        For i = 1 To deviceCount
            device.SetId(deviceIds(i))
            devName = device.GetName()

            If regEx.Test(devName) Then
                ' ��� -UZ ��������� �� ��������� �������� ����� ���������� �� "�������"
                If Not deviceInfo.Exists(devName) Then
                    deviceInfo.Add devName, device.GetComponentName()
                    global_e3App.PutInfo 0, "������ ������ -UZ: " & devName & _
                        ", ������� ���������: '" & device.GetComponentName() & "'"
                End If
            End If
        Next
    Else
        global_e3App.PutInfo 1, "� ������� �� ������� ���������."
    End If

    If deviceInfo.Count = 0 Then
        global_e3App.PutInfo 1, "��������: ���������� -UZ, ��������������� �������, �� �������."
    End If

    global_e3App.PutInfo 0, "����� ��������� -UZ ��������."
    Set deviceInfo = Nothing
    Call CleanUpObjects(regEx, Nothing, device, Nothing, Nothing)
End Sub

'*******************************************************************************
' ��������� UpdateUZComponents(job, e3App_local)
' ��������� ��� ���������� ��� ��������� -UZ �� ������ ������� ������������.
'*******************************************************************************
Sub UpdateUZComponents(job, e3App_local)
    Dim deviceIds, deviceCount, i, OOSIndex
    Dim device, devName, currentComponentName, newComponentName, extractedArticle
    Dim regEx, matches
    Dim targetUZName
    Dim componentVersion ' ��������� ���������� ��� ������ ����������

    ' ������������� ������ ����������, ��� �� �������
    componentVersion = "1" 

    Set device = job.CreateDeviceObject()
    Set regEx = New RegExp
    regEx.IgnoreCase = True

    e3App_local.PutInfo 0, "�������� ���������� ����������� ��� ��������� -UZ..."

    ' ��������� �� ���� ����������� ��������� �� OOS ��������
    For Each OOSIndex In global_OOSArticles.Keys()
        extractedArticle = global_OOSArticles.Item(OOSIndex)
        targetUZName = "-UZ" & OOSIndex ' ��������� ��� �������� -UZ ����������

        e3App_local.PutInfo 0, "������������ OOS-�������: '" & extractedArticle & "' ��� �������������� -UZ ����������: '" & targetUZName & "'"

        ' ���� ������� -UZ ���������� � �������
        deviceCount = job.GetAllDeviceIds(deviceIds)
        Dim foundTargetDevice : foundTargetDevice = False

        If deviceCount > 0 Then
            For i = 1 To deviceCount
                device.SetId(deviceIds(i))
                devName = device.GetName()

                ' ���������, ������������� �� ��� ���������� �������� -UZ
                If LCase(devName) = LCase(targetUZName) Then
                    currentComponentName = device.GetComponentName()
                    e3App_local.PutInfo 0, "������ ������: " & devName & ", ������� ���������: '" & currentComponentName & "'"

                    ' ���� ����� ��� ���������� � global_articleComponentMap
                    If global_articleComponentMap.Exists(extractedArticle) Then
                        newComponentName = global_articleComponentMap.Item(extractedArticle)

                        If LCase(currentComponentName) <> LCase(newComponentName) Then
                            On Error Resume Next ' �������� ��������� ������ ��� SetComponentName
                            ' �����������: �������� ��� ���������� � ������
                            device.SetComponentName newComponentName, componentVersion
                            If Err.Number = 0 Then
                                e3App_local.PutInfo 0, "�������: �������� ��������� ������� '" & devName & _
                                    "' � '" & currentComponentName & "' �� '" & newComponentName & "' (������: " & componentVersion & ")"
                            Else
                                e3App_local.PutInfo 2, "������: �� ������� �������� ��������� ������� '" & devName & _
                                    "' �� '" & newComponentName & "'. ������: " & Err.Description & _
                                    " (���: " & Err.Number & ", ��������: " & Err.Source & ")" ' ������� Err.Source ��� ������ �����������
                            End If
                            On Error GoTo 0 ' ��������� ��������� ������
                        Else
                            e3App_local.PutInfo 0, "������ '" & devName & "' ��� ����� ��������� ���������: '" & newComponentName & "'. ����������."
                        End If
                    Else
                        e3App_local.PutInfo 1, "��������������: ������� '" & extractedArticle & _
                            "' ��� ������� '" & devName & "' �� ������ � ������� ������������ global_articleComponentMap. ��������� �� ��������."
                    End If
                    foundTargetDevice = True
                    Exit For ' ������ ������ � ���������, ��������� � ���������� OOS-��������
                End If
            Next
        End If

        If Not foundTargetDevice Then
            e3App_local.PutInfo 1, "��������������: ������� ������ '" & targetUZName & "' �� ������ � ������� ��� �������� '" & extractedArticle & "'. ����������."
        End If

    Next

    e3App_local.PutInfo 0, "���������� ����������� ��� ��������� -UZ ���������."
    Call CleanUpObjects(regEx, Nothing, device, Nothing, Nothing)
End Sub

'*******************************************************************************
' �������� ���� ���������� �������
'*******************************************************************************
Sub Main()

    ' 1. ������������� �������� E3.series
    Set global_e3App = CreateObject("CT.Application")
    If global_e3App Is Nothing Then
        MsgBox "�� ������� ������������ � E3.series. ���������, ��� E3.series �������.", vbCritical, "������ E3.series"
        Exit Sub
    End If
    Set global_job = global_e3App.CreateJobObject()
    If global_job Is Nothing Then
        MsgBox "�� ������� �������� ������ � �������� ������� E3.series.", vbCritical, "������ E3.series"
        Call CleanUpObjects(Nothing, Nothing, Nothing, Nothing, global_e3App)
        Exit Sub
    End If

    ' ������� ���� ��������� E3.series � ������ ����������
    global_e3App.PutMessageEx 0, "������ ������� E3_UZ_ComponentUpdater...", 0, 0, 0, 249 ' ���� BLUE

    ' 2. ������������� ���������� ��������
    Set global_OOSArticles = CreateObject("Scripting.Dictionary")
    Set global_articleComponentMap = CreateObject("Scripting.Dictionary")

    ' 3. ���������� global_articleComponentMap
    ' ���� ������ ����� �������� ������ ������� �����.
    ' ������: global_articleComponentMap.Add "���_�������", "���_�����_���������"
    global_articleComponentMap.Add "ABA00005", "VF51_0.75"
    global_articleComponentMap.Add "ABA00006", "VF51_1.5"
    global_articleComponentMap.Add "ABA00011", "VF51_11.0"
    global_articleComponentMap.Add "ABA00012", "VF51_15.0"
    global_articleComponentMap.Add "ABA00013", "VF51_18.5"
    global_articleComponentMap.Add "ABA00007", "VF51_2.2"
    global_articleComponentMap.Add "ABA00014", "VF51_22.0"
    global_articleComponentMap.Add "ABA00107", "VF51_3.0"
    global_articleComponentMap.Add "ABA00008", "VF51_4.0"
    global_articleComponentMap.Add "ABA00009", "VF51_5.5"
    global_articleComponentMap.Add "ABA00010", "VF51_7.5"
    global_articleComponentMap.Add "ABC00123", "VF101_0.75"
    global_articleComponentMap.Add "ABC00124", "VF101_1.5"
    global_articleComponentMap.Add "ABC00129", "VF101_11"
    global_articleComponentMap.Add "ABC00130", "VF101_15"
    global_articleComponentMap.Add "ABC00131", "VF101_18.5"
    global_articleComponentMap.Add "ABC00125", "VF101_2.2"
    global_articleComponentMap.Add "ABC00132", "VF101_22"
    global_articleComponentMap.Add "ABC00133", "VF101_30"
    global_articleComponentMap.Add "ABC00134", "VF101_37"
    global_articleComponentMap.Add "ABC00135", "VF101_45"
    global_articleComponentMap.Add "ABC00127", "VF101_5.5"
    global_articleComponentMap.Add "ABC00166", "VF101_55"
    global_articleComponentMap.Add "ABC00128", "VF101_7.5"
    global_articleComponentMap.Add "ABC00137", "VF101_75"
    global_articleComponentMap.Add "ABC00138", "VF101_90"
    global_articleComponentMap.Add "ABC00126", "VF101_4"

    ' 4. ���������� �������� ��������
    Call ExtractOOSArticles()        ' ��������� �������� �� OOS ��������
    Call FindUZDevices(global_job)   ' ���� ���������� -UZ (��� �����������/�����������)
    Call UpdateUZComponents(global_job, global_e3App) ' ��������� ���������� -UZ ���������

    global_e3App.PutInfo 0, "������ E3_UZ_ComponentUpdater ��������."

    ' 5. ������� ���������� ��������
    Call CleanUpObjects(Nothing, Nothing, Nothing, global_job, global_e3App)
    Set global_OOSArticles = Nothing
    Set global_articleComponentMap = Nothing

    Exit Sub ' ����� �� Sub Main ��� �������� ����������

ErrorHandler:
    global_e3App.PutInfo 2, "����������� ������ � �������: " & Err.Description & " (���: " & Err.Number & ")"
    MsgBox "��������� ����������� ������ � �������. ��. ���� ��������� E3.series ��� ������������.", vbCritical, "������ �������"
    Call CleanUpObjects(Nothing, Nothing, Nothing, global_job, global_e3App)
    Set global_OOSArticles = Nothing
    Set global_articleComponentMap = Nothing
End Sub

' ������ �������� ���������
Call Main()