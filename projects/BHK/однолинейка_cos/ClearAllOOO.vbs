Option Explicit

' === ������� === ���������� ������ �� ����� �������
Function ExtractNumber(ByVal itemName)
    Dim re, matches
    Set re = New RegExp
    ' ���� ����� � ����� ������ ����� �������� (��������, OOS)
    re.Pattern = "(\d+)$"
    re.Global = False
    
    Set matches = re.Execute(itemName)
    
    If matches.Count > 0 Then
        ExtractNumber = CInt(matches.Item(0).Value)
    Else
        ExtractNumber = 0 ' ���� ����� �� ������
    End If
    
    Set re = Nothing
End Function

' === ��������� === ����� ���� �������� OOS � �������
Sub FindAllOOSSymbols(ByRef OOSSymbols)
    Dim e3App, job, symbol
    Dim symbolIds(), symbolCount
    Dim i, symbolName, symbolNumber
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    
    e3App.PutInfo 0, "=== ����� ���� �������� OOS � ������� ==="
    
    symbolCount = job.GetSymbolIds(symbolIds)
    If symbolCount = 0 Then
        e3App.PutInfo 0, "� ������� �� ������� ��������."
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    
    For i = 1 To symbolCount
        symbol.SetId(symbolIds(i))
        symbolName = symbol.GetName()
        
        If LCase(Left(symbolName, 3)) = "OOS" Then
            symbolNumber = ExtractNumber(symbolName)
            If symbolNumber > 0 Then
                OOSSymbols.Add symbolNumber, symbolIds(i)
                e3App.PutInfo 0, "������ ������ OOS: " & symbolName & " (�����: " & symbolNumber & ", ID: " & symbolIds(i) & ")"
            Else
                e3App.PutInfo 0, "������ OOS ������, �� ����� �� ���������: " & symbolName
            End If
        End If
    Next
    
    e3App.PutInfo 0, "����� ������� �������� OOS � ��������: " & OOSSymbols.Count
    
    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === ��������� === ������� ��������� ������� OOS
Sub ClearOOSSymbolAttributes(ByVal OOSSymbolId, ByVal number)
    Dim e3App, job, symbol
    
    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    
    symbol.SetId(OOSSymbolId)
    
    e3App.PutInfo 0, "=== ������� ��������� ������� OOS" & number & " ==="
    
    ' ������� ��������� QF ����������
    symbol.SetAttributeValue "�� V_Inom", "-"
    e3App.PutInfo 0, "������ ������� �� V_Inom"
    
    symbol.SetAttributeValue "�� V_Type", "-"
    e3App.PutInfo 0, "������ ������� �� V_Type"
    
    symbol.SetAttributeValue "�� V_Icu", "-"
    e3App.PutInfo 0, "������ ������� �� V_Icu"
    
    symbol.SetAttributeValue "�� V_Proizv", "-"
    e3App.PutInfo 0, "������ ������� �� V_Proizv"
    
    symbol.SetAttributeValue "�� V_Dop ystr", "-"
    e3App.PutInfo 0, "������ ������� �� V_Dop ystr"
    
    ' ������� ��������� KM ����������
    symbol.SetAttributeValue "�� K_Type", "-"
    e3App.PutInfo 0, "������ ������� �� K_Type"
    
    symbol.SetAttributeValue "�� K_Proizv", "-"
    e3App.PutInfo 0, "������ ������� �� K_Proizv"
    
    symbol.SetAttributeValue "�� K_Inom", "-"
    e3App.PutInfo 0, "������ ������� �� K_Inom"
    
    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' === �������� ��������� === ������� ��������� ���� �������� OOS
Sub ClearAllOOSSymbolsAttributes()
    ' ���������� ��������������� ����
    Dim msgResult
    msgResult = MsgBox("�������� ������ ���������?", vbOKCancel + vbQuestion, "�������������")
    
    ' ���� ������������ ����� "������", ������� �� �������
    If msgResult = vbCancel Then
        Exit Sub
    End If
    
    Dim e3App
    Dim OOSSymbols
    Dim OOSNumber, OOSSymbolId
    
    Set e3App = CreateObject("CT.Application")
    Set OOSSymbols = CreateObject("Scripting.Dictionary")
    
    e3App.PutInfo 0, "=== ����� ������� ��������� ���� OOS �������� ==="
    
    ' ������� ��� ������� OOS
    Call FindAllOOSSymbols(OOSSymbols)
    
    If OOSSymbols.Count = 0 Then
        e3App.PutInfo 0, "������� OOS �� �������. ������� �� ���������."
        Set OOSSymbols = Nothing
        Set e3App = Nothing
        Exit Sub
    End If
    
    ' ������� �������� ������� ������� OOS
    For Each OOSNumber In OOSSymbols.Keys
        OOSSymbolId = OOSSymbols.Item(OOSNumber)
        
        e3App.PutInfo 0, "--- ������� OOS" & OOSNumber & " ---"
        
        Call ClearOOSSymbolAttributes(OOSSymbolId, OOSNumber)
    Next
    
    e3App.PutInfo 0, "=== ���������� ������� ���� OOS �������� ==="
    e3App.PutInfo 0, "���������� ��������: " & OOSSymbols.Count
    
    Set OOSSymbols = Nothing
    Set e3App = Nothing
End Sub

' === �������� ������ ===
Dim e3App
Set e3App = CreateObject("CT.Application")

e3App.PutInfo 0, "=== ����� ������� ������� ��������� OOS �������� ==="
Call ClearAllOOSSymbolsAttributes()
e3App.PutInfo 0, "=== ����� ������� ==="

Set e3App = Nothing