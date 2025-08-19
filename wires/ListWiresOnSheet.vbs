                    App.PutInfo 0, "������� '����� ���������� (����������)' ���������� ��� coreID: " & coreIds(j)
Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject()
Set Device = Job.CreateDeviceObject()
Set Sheet = Job.CreateSheetObject()
Set Pin = Job.CreatePinObject()

' ��������� �������
Dim sheetIds()
Dim netIds()
Dim coreIds()

' ������ ������ tree ��� ������ � ������� �������
Dim Tree
Set Tree = Job.CreateTreeObject
' ������������� �������� ������ �������
TreeId = Tree.SetId(Job.GetActiveTreeId())

' ����� ���������� � ������ ������
Dim sheetCount : sheetCount = Tree.GetSelectedSheetIds(sheetIds)

If sheetCount = 0 Then
    ' ���� ��� ���������� ������, ����� ��������
    sheetCount = 1
    ReDim sheetIds(1)
    sheetIds(1) = Job.GetActiveSheetId()
    
    If sheetIds(1) = 0 Then
        App.PutInfo 1, "��� ��������� ��� ����������� �����"
        WScript.Quit
    End If
    
    App.PutInfo 0, "�������� � �������� ������..."
Else
    App.PutInfo 0, "�������� � ����������� �������..."
End If

' ���������� ��� ��������� �����
For selectedNum = 1 To sheetCount
    ' ������������� ������� ����
    Sheet.SetId sheetIds(selectedNum)
    Dim selectedSheetName : selectedSheetName = Sheet.GetName()
    App.PutInfo 0, "��������� �����: " & selectedSheetName

' �������� ���� �� �����
Dim netCount : netCount = Sheet.GetNetIds(netIds)
If netCount = 0 Then ReDim netIds(0)
App.PutInfo 0, "=============== ������� �� ����� " & selectedSheetName & " ==============="

If netCount = 0 Then
    App.PutInfo 1, "�� ����� ��� �����"
Else
    ' ������� ������ ��� ������ � ������
    Dim Net : Set Net = Job.CreateNetObject()
    Dim wireCount : wireCount = 0

    ' ���������� ��� ����
    For i = 1 To netCount
        Net.SetId netIds(i)
        
        ' �������� ��� ���� � ����
        Dim coreCount : coreCount = Net.GetCoreIds(coreIds)
        If coreCount = 0 Then ReDim coreIds(0)
        If coreCount > 0 Then
            ' ���������� ����
            For j = 1 To coreCount
                Pin.SetId coreIds(j)
                Device.SetId coreIds(j)
                
                ' ���� ��� ������
                If Device.IsWiregroup() Then
                    wireCount = wireCount + 1
                    ' �������� ���������� � �������
                    Dim wireName : wireName = Device.GetName()
                    Dim wireId : wireId = Device.GetId()
                    Dim signalName : signalName = Pin.GetSignalName()
                    Dim colorDesc : colorDesc = Pin.GetColourDescription()
                    
                    ' ������� ���������� � ������������ ID
                    App.PutMessageEx 0, wireCount & ". " & " (����: " & signalName & ", ����: " & colorDesc & ", CoreID: " & coreIds(j) & ")", coreIds(j), 0, 0, 0
                        ' ������������� ������� "����� ���������� (����������)" ��� core
                        Pin.SetAttributeValue "����� ���������� (����������)", "230/400V"
                End If
            Next
        End If
    Next

    ' �������
    Set Net = Nothing

    ' ������� ����
    If wireCount = 0 Then
        App.PutInfo 0, "�� ����� �� ������� ��������"
    Else
        App.PutInfo 0, "==============================================="
        App.PutInfo 0, "����� ������� �������� �� ����� " & selectedSheetName & ": " & wireCount
    End If
End If
Next ' ����� ����� �� ������

' ������� �������
Set Pin = Nothing
Set Sheet = Nothing
Set Device = Nothing
Set Job = Nothing
Set App = Nothing
