Option Explicit

' ���������� ���������� ��������
Dim global_OOSArticles       ' ��� �������� �������� OOS � ����������� ���������
Dim global_articleComponentMap ' ��� �������� ������������ ������� -> ��� ���������� ��������

' === ��������������� ������� === ������������ ����� ���������� (�������������� ������� ������������� �������� � ���������/��������)
' ��� ������� ��������� � ����, �� �� ����� �������������� ��� ���������� global_articleComponentMap
' �� ������� ������������. ��� ����� ���� ������� ��� ������ �����.
Function NormalizeComponentName(ByVal compName)
    Dim newCompName : newCompName = compName

    ' ������ ������� ������������� �������� �� ���������/�������� �������
    ' ��� ������� �������� ������ ��-�� ������� � ��������� (��������, ������������� � vs ��������� P)
    newCompName = Replace(newCompName, "�", "A") ' Cyrillic A -> Latin A
    newCompName = Replace(newCompName, "�", "B") ' Cyrillic Ve -> Latin B
    newCompName = Replace(newCompName, "�", "E") ' Cyrillic Ye -> Latin E
    newCompName = Replace(newCompName, "�", "K") ' Cyrillic Ka -> Latin K
    newCompName = Replace(newCompName, "�", "M") ' Cyrillic Em -> Latin M
    newCompName = Replace(newCompName, "�", "H") ' Cyrillic En -> Latin H
    newCompName = Replace(newCompName, "�", "O") ' Cyrillic O -> Latin O
    newCompName = Replace(newCompName, "�", "P") ' Cyrillic Er -> Latin P (�������� ����� ��� "3�" -> "3P")
    newCompName = Replace(newCompName, "�", "C") ' Cyrillic Es -> Latin C
    newCompName = Replace(newCompName, "�", "T") ' Cyrillic Te -> Latin T
    newCompName = Replace(newCompName, "�", "X") ' Cyrillic Kha -> Latin X
    newCompName = Replace(newCompName, "�", "Y") ' Cyrillic U -> Latin Y
    
    ' ����������� ������ ��� "��" -> "3P" (���� "�" - ��� �������� ��� "3" � ���������� � "�")
    newCompName = Replace(newCompName, "��", "3P") 
    ' ����� �������� ������ � �� 3
    newCompName = Replace(newCompName, "�", "3")


    NormalizeComponentName = newCompName
End Function


' === ������� ��������� === ���������� ��������� �� ��������� OOS ��������
Sub ExtractOOSArticles()
    Dim e3App, job, symbol
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName
    Dim s
    
    ' ���������� ��� ���������
    Dim dProizv3Value, dProizv2Value
    Dim extractedArticle ' ��� �������� ������������ ��������

    ' ��� ����������� ���������
    Dim re, matches
    Set re = New RegExp
    re.Pattern = "(ABA|ABC)\d+" 
    re.IgnoreCase = True  
    re.Global = False     

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()

    e3App.PutInfo 0, "=== ����� �������: ���������� ��������� OOS �������� ==="

    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "� ������� ��� �������� ��� �������. ������ ��������."
        Call CleanUpObjects(re, matches, symbol, job, e3App) 
        Exit Sub
    End If

    e3App.PutInfo 0, "������� " & allSymbolCount & " �������� � �������. ����� OOS �������� �� �� D_Proizv3=1..."
    
    Dim foundOOSCount : foundOOSCount = 0

    For s = 1 To allSymbolCount 
        currentSymbolId = allSymbolIds(s)
        symbol.SetId(currentSymbolId)
        symbolName = symbol.GetName()

        If LCase(Left(symbolName, 3)) = "OOS" Then
            dProizv3Value = Trim(CStr(symbol.GetAttributeValue("�� D_Proizv3")))

            If dProizv3Value = "1" Then
                foundOOSCount = foundOOSCount + 1
                e3App.PutInfo 0, "  ������ OOS ������: '" & symbolName & "' (ID: " & currentSymbolId & ") � �� D_Proizv3 = 1."

                dProizv2Value = Trim(CStr(symbol.GetAttributeValue("�� D_Proizv2")))
                e3App.PutInfo 0, "    �������� �� D_Proizv2: '" & dProizv2Value & "'"

                Dim OOSIndex
                On Error Resume Next 
                OOSIndex = CLng(Mid(symbolName, 4))
                If Err.Number <> 0 Then
                    OOSIndex = "����������" 
                    Err.Clear
                End If
                On Error GoTo 0

                If Len(dProizv2Value) > 0 Then
                    Set matches = re.Execute(dProizv2Value)
                    If matches.Count > 0 Then
                        extractedArticle = matches.Item(0).Value 
                        e3App.PutInfo 0, "    ����������� �������: '" & extractedArticle & "'"
                        
                        If Not global_OOSArticles.Exists(CStr(OOSIndex)) Then
                            global_OOSArticles.Add CStr(OOSIndex), extractedArticle
                        Else
                            global_OOSArticles.Item(CStr(OOSIndex)) = extractedArticle 
                        End If
                    Else
                        extractedArticle = "�� �������"
                        e3App.PutInfo 0, "    ������� �� ������ �� ������� � �� D_Proizv2."
                    End If
                Else
                    extractedArticle = "�����"
                    e3App.PutInfo 0, "    �� D_Proizv2 ����, ������� �� ����� ���� ��������."
                End If
                
                e3App.PutInfo 0, "  --- ��������� ��� OOS" & OOSIndex & ": ������� = " & extractedArticle & " ---"
            End If
        End If
    Next 

    If foundOOSCount = 0 Then
        e3App.PutInfo 0, "�� ������� OOS �������� �� ��������� �������� �� D_Proizv3 ������ '1'."
    End If

    e3App.PutInfo 0, "=== ���������� ������� ==="

    Call CleanUpObjects(re, matches, symbol, job, e3App)
End Sub

' === ��������������� ��������� === ������� ��������
Sub CleanUpObjects(reObj, matchesObj, symbolObj, jobObj, appObj)
    Set reObj = Nothing
    Set matchesObj = Nothing
    Set symbolObj = Nothing
    Set jobObj = Nothing
    Set appObj = Nothing
End Sub

' === ��������� FindQFKM (�� ��������, ��������� ��� �������) ===
Sub FindQFKM(job)
    Dim device, re, matches
    Dim deviceIds, deviceCount
    Dim i, j
    Dim devName, compName
    Dim deviceInfo
    Dim key, arr, k
    Dim e3App_local 

    Set e3App_local = CreateObject("CT.Application") 

    Set device = job.CreateDeviceObject()
    Set re = New RegExp
    re.Pattern = "^-(QF|KM)\d+$" 
    re.IgnoreCase = False
    re.Global = False

    deviceCount = job.GetAllDeviceIds(deviceIds)

    If deviceCount = 0 Then
        e3App_local.PutInfo 0, "���������� QF/KM �� ������� � �������." 
        Set device = Nothing
        Set re = Nothing
        Set e3App_local = Nothing 
        Exit Sub
    End If

    Set deviceInfo = CreateObject("Scripting.Dictionary")

    e3App_local.PutInfo 0, "=== FindQFKM: ������ ������ ��������� -QF � ����������� '�������' � ���� -KM ��������� ===" 

    For i = 1 To deviceCount
        device.SetId(deviceIds(i))
        devName = device.GetName()

        If re.Test(devName) Then 
            compName = device.GetComponentName()
            
            If LCase(Left(devName, 3)) = "-qf" Then 
                If InStr(1, LCase(compName), "�������") > 0 Then 
                    If Not deviceInfo.Exists(devName) Then
                        deviceInfo.Add devName, Array() 
                    End If
                    Dim oldArr_qf
                    oldArr_qf = deviceInfo.Item(devName) 
                    Dim currentSize_qf : currentSize_qf = -1
                    On Error Resume Next 
                    currentSize_qf = UBound(oldArr_qf)
                    On Error GoTo 0 
                    Dim newArr_qf
                    If currentSize_qf = -1 Then 
                        ReDim newArr_qf(0) 
                    Else
                        ReDim Preserve newArr_qf(currentSize_qf + 1)
                    End If
                    For j = LBound(oldArr_qf) To currentSize_qf 
                        newArr_qf(j) = oldArr_qf(j)
                    Next
                    newArr_qf(UBound(newArr_qf)) = "ID=" & deviceIds(i) & ", Component=" & compName
                    deviceInfo.Item(devName) = newArr_qf
                Else
                    e3App_local.PutInfo 0, "  ��������� ����������: " & devName & " (ID: " & deviceIds(i) & ") - ��������� '" & compName & "' �� �������� '�������'."
                End If
            ElseIf LCase(Left(devName, 3)) = "-km" Then 
                If Not deviceInfo.Exists(devName) Then
                    deviceInfo.Add devName, Array() 
                End If
                Dim oldArr_km
                oldArr_km = deviceInfo.Item(devName) 
                Dim currentSize_km : currentSize_km = -1
                On Error Resume Next 
                currentSize_km = UBound(oldArr_km)
                On Error GoTo 0 
                Dim newArr_km
                If currentSize_km = -1 Then 
                    ReDim newArr_km(0) 
                Else
                    ReDim Preserve newArr_km(currentSize_km + 1)
                End If
                For j = LBound(oldArr_km) To currentSize_km 
                    newArr_km(j) = oldArr_km(j)
                Next
                newArr_km(UBound(newArr_km)) = "ID=" & deviceIds(i) & ", Component=" & compName
                deviceInfo.Item(devName) = newArr_km
            End If 
        End If 
    Next

    e3App_local.PutInfo 0, "=== FindQFKM: ������� " & deviceInfo.Count & " ����� ��������� ===" 

    If deviceInfo.Count > 0 Then
        For Each key In deviceInfo.Keys
            e3App_local.PutInfo 0, "����������: " & key 
            arr = deviceInfo.Item(key)
            For k = LBound(arr) To UBound(arr)
                e3App_local.PutInfo 0, "    " & arr(k) 
            Next
        Next
    Else
        e3App_local.PutInfo 0, "�� ������� ���������, ��������������� ������� -QF### (� '�������' � ����������) ��� -KM###." 
    End If

    Set device = Nothing
    Set re = Nothing
    Set deviceInfo = Nothing
    Set e3App_local = Nothing 
End Sub


' === ����� ��������� === ���������� ����� ���������� QF �� ������ �������� OOS
Sub COMM(job, e3App_local)
    Dim device
    Dim OOSIndex, extractedArticle
    Dim targetDeviceName, targetDeviceId
    Dim currentDeviceName, currentComponentName
    Dim allDeviceIds(), deviceCount
    Dim i
    Dim newComponentName 
    Dim componentVersion 

    Set device = job.CreateDeviceObject()

    e3App_local.PutInfo 0, "=== COMM: ������ ��������� ���������� ����������� ==="

    If global_OOSArticles.Count = 0 Then
        e3App_local.PutInfo 0, "COMM: ��� ������ �� OOS �������� � �� D_Proizv3=1 ��� ���������."
        Set device = Nothing 
        Exit Sub
    End If

    For Each OOSIndex In global_OOSArticles.Keys 
        extractedArticle = global_OOSArticles.Item(OOSIndex) 
        
        targetDeviceName = "-QF" & OOSIndex 
        e3App_local.PutInfo 0, "  ����� ���������� '" & targetDeviceName & "' ��� ���������� (OOS" & OOSIndex & ", �������: " & extractedArticle & ")..."

        targetDeviceId = 0 
        deviceCount = job.GetAllDeviceIds(allDeviceIds) 
        
        For i = 1 To deviceCount
            device.SetId(allDeviceIds(i))
            currentDeviceName = device.GetName()
            currentComponentName = device.GetComponentName()

            If UCase(currentDeviceName) = UCase(targetDeviceName) Then
                If InStr(1, LCase(currentComponentName), "�������") > 0 Then
                    targetDeviceId = allDeviceIds(i)
                    e3App_local.PutInfo 0, "    ������� ���������� '" & targetDeviceName & "' (ID: " & targetDeviceId & ") � ����������� '�������'."
                    Exit For 
                Else
                    e3App_local.PutInfo 0, "    ���������� '" & targetDeviceName & "' �������, �� ��� ��������� ('" & currentComponentName & "') �� �������� '�������'. ����������."
                End If
            End If
        Next

        If targetDeviceId > 0 Then
            device.SetId(targetDeviceId)
            
            If global_articleComponentMap.Exists(extractedArticle) Then
                ' �������� ��� ���������� �� ������� (��� ������������ �����, ��� ��� ��� ��� ������� ��� ����������)
                newComponentName = global_articleComponentMap.Item(extractedArticle) 
                componentVersion = "1" 
                
                On Error Resume Next 
                device.SetComponentName newComponentName, componentVersion 
                If Err.Number = 0 Then
                    e3App_local.PutInfo 0, "    �������: ��������� ���������� '" & targetDeviceName & "' �������� ��: '" & newComponentName & "' (������: '" & componentVersion & "')."
                Else
                    e3App_local.PutInfo 0, "    ������ ��� ���������� ���������� ��� '" & targetDeviceName & "': " & Err.Description
                    Err.Clear 
                End If
                On Error GoTo 0 
            Else
                e3App_local.PutInfo 0, "    ��������������: ��� �������� '" & extractedArticle & "' (OOS" & OOSIndex & ") �� ������� ������������ � ������� �����������. ��������� �� ��������."
            End If
        Else
            e3App_local.PutInfo 0, "    ��������������: ���������� '" & targetDeviceName & "' (� ����������� '�������') �� ������� � ������� ��� OOS" & OOSIndex & ". ���������� ���������."
        End If
    Next

    e3App_local.PutInfo 0, "=== COMM: ���������� ��������� ���������� ����������� ==="

    Set device = Nothing 
End Sub

' === �������� ������ ===
Dim global_e3App, global_job
Set global_e3App = CreateObject("CT.Application")
Set global_job = global_e3App.CreateJobObject()

Set global_OOSArticles = CreateObject("Scripting.Dictionary") 
Set global_articleComponentMap = CreateObject("Scripting.Dictionary") 

' ���������� global_articleComponentMap ������� �� ��������������� ��������� �������,
' ������� image_6c9280.png, image_917c55.png, image_917855.png � ����� ������ �� ���������� �������.
' ��� ���������� NormalizeComponentName � ��������� (������ �����������), �� ������� ������������.
global_articleComponentMap.Add "ABA00002", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABA00110", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABA00005", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABA00104", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABA00003", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABA00006", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABA00105", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABA00011", "�������_3�_50A_13182DEK"
global_articleComponentMap.Add "ABA00111", "�������_3�_50A_13182DEK"
global_articleComponentMap.Add "ABA00012", "�������_3�_50A_13182DEK"
global_articleComponentMap.Add "ABA00112", "�������_3�_50A_13182DEK"
global_articleComponentMap.Add "ABA00013", "�������_3�_63A_13183DEK"
global_articleComponentMap.Add "ABA00113", "�������_3�_63A_13183DEK"
global_articleComponentMap.Add "ABA00004", "�������_3�_16A_13177DEK"
global_articleComponentMap.Add "ABA00007", "�������_3�_16A_13177DEK"
global_articleComponentMap.Add "ABA00106", "�������_3�_16A_13177DEK"
global_articleComponentMap.Add "ABA00014", "�������_3�_80A_13008DEK"
global_articleComponentMap.Add "ABA00114", "�������_3�_80A_13008DEK"
global_articleComponentMap.Add "ABA00107", "�������_3�_20A_13178DEK"
global_articleComponentMap.Add "ABA00008", "�������_3�_20A_13178DEK"
global_articleComponentMap.Add "ABA00108", "�������_3�_20A_13178DEK"
global_articleComponentMap.Add "ABA00009", "�������_3�_25A_13179DEK"
global_articleComponentMap.Add "ABA00109", "�������_3�_25A_13179DEK"
global_articleComponentMap.Add "ABA00010", "�������_3�_25A_13179DEK"
global_articleComponentMap.Add "ABC00023", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABC00024", "�������_3�_10A_13176DEK"
global_articleComponentMap.Add "ABC00029", "�������_3�_40A_13181DEK"
global_articleComponentMap.Add "ABC00030", "�������_3�_50A_13182DEK"
global_articleComponentMap.Add "ABC00031", "�������_3�_63A_13183DEK"
global_articleComponentMap.Add "ABC00025", "�������_3�_16A_13177DEK"
global_articleComponentMap.Add "ABC00032", "�������_3�_80A_13008DEK"
global_articleComponentMap.Add "ABC00033", "�������_3�_100A_13009DEK"
global_articleComponentMap.Add "ABC00034", "�������_3�_125A_13027DEK"
global_articleComponentMap.Add "ABC00035", "�������_3�_160A_22752DEK"
global_articleComponentMap.Add "ABC00027", "�������_3�_25A_13179DEK"
global_articleComponentMap.Add "ABC00066", "�������_3�_200A_22754DEK"
global_articleComponentMap.Add "ABC00028", "�������_3�_32A_13180DEK"
global_articleComponentMap.Add "ABC00037", "�������_3�_200A_22754DEK"
global_articleComponentMap.Add "ABC00067", "�������_3�_200A_22754DEK"
global_articleComponentMap.Add "ABC00038", "�������_3�_250A_22756DEK"
global_articleComponentMap.Add "ABC00068", "�������_3�_250A_22756DEK"
global_articleComponentMap.Add "ABC00026", "�������_3�_20A_13178DEK"
global_articleComponentMap.Add "ABA00001", "�������_3�_40A_13181DEK"
global_articleComponentMap.Add "ABC00036", "�������_3�_160A_22752DEK"
global_articleComponentMap.Add "ABC00123", "�������_3�_10A_13176DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00124", "�������_3�_10A_13176DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00129", "�������_3�_40A_13181DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00130", "�������_3�_50A_13182DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00131", "�������_3�_63A_13183DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00125", "�������_3�_16A_13177DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00132", "�������_3�_80A_13008DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00133", "�������_3�_100A_13009DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00134", "�������_3�_125A_13027DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00135", "�������_3�_160A_22752DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00127", "�������_3�_25A_13179DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00166", "�������_3�_200A_22754DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00128", "�������_3�_32A_13180DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00137", "�������_3�_200A_22754DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00167", "�������_3�_200A_22754DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00138", "�������_3�_250A_22756DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00168", "�������_3�_250A_22756DEK" ' <-- ���������
global_articleComponentMap.Add "ABC00126", "�������_3�_20A_13178DEK" ' <-- ���������


' �������� �������� ���������
Call ExtractOOSArticles() 
Call FindQFKM(global_job) 
Call COMM(global_job, global_e3App) 

' ��������� ������� ���������� ��������
Set global_job = Nothing
Set global_e3App = Nothing
Set global_OOSArticles = Nothing 
Set global_articleComponentMap = Nothing 
