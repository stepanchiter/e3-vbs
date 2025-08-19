Option Explicit

' === ������� ��������� === �������������� �������� OOS �� �������
Sub RenameOOSSymbolsByLocation()
    Dim e3App, job, symbol, sheet
    Dim allSymbolIds(), allSymbolCount
    Dim currentSymbolId, symbolName
    Dim s, i, j
    Dim OOSCounter : OOSCounter = 0 ' ������� ��� ����������������� ����������

    ' ���������� ��� �������� ������ OOS �������� ��� ����������
    ' ������ ������� � OOSSymbolsToRename ����� ��������:
    ' (SymbolID, SheetID, SheetName, Column, Row, X, Y, OriginalName)
    Dim OOSSymbolsToRename()
    ReDim OOSSymbolsToRename(0) ' ������������� � ��������� ���������, ����� ���������������

    Dim OOSCountPlaced : OOSCountPlaced = 0 ' ������� ������ ��� ����������� OOS ��������

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()
    Set sheet = job.CreateSheetObject() ' ��� ��������� ����� �����

    e3App.PutInfo 0, "=== ����� �������: �������������� OOS �������� �� ������� ==="

    ' �������� ��� ID �������� � �������
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "� ������� ��� �������� ��� �������. ������ ��������."
        Set symbol = Nothing
        Set sheet = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    e3App.PutInfo 0, "������� " & allSymbolCount & " �������� � �������. �������� ������ � �������� ��������������� OOS �������..."

    ' === ������ ������: ���� ������ � ����������� OOS �������� � ��������� �������������� ===
    For s = 1 To allSymbolCount
        currentSymbolId = allSymbolIds(s)
        symbol.SetId(currentSymbolId)
        symbolName = symbol.GetName()

        If LCase(Left(symbolName, 3)) = "oos" Then
            Dim xPos, yPos, gridDesc, columnValue, rowValue
            Dim sheetId : sheetId = symbol.GetSchemaLocation(xPos, yPos, gridDesc, columnValue, rowValue)

            If sheetId > 0 Then ' ������ �������� �� �����
                OOSCountPlaced = OOSCountPlaced + 1
                ReDim Preserve OOSSymbolsToRename(OOSCountPlaced) ' ����������� ������ �������

                sheet.SetId sheetId
                Dim sheetName : sheetName = sheet.GetName()

                ' Store the data for sorting, including original name
                OOSSymbolsToRename(OOSCountPlaced) = Array(currentSymbolId, sheetId, sheetName, columnValue, rowValue, xPos, yPos, symbolName)
                e3App.PutInfo 0, "  OOS ������ ������ �� �����: " & symbolName & " (ID: " & currentSymbolId & ") �� �����: " & sheetName & " " & columnValue & rowValue & " (" & xPos & ", " & yPos & ")"
                
                ' �������� ��������������� ������, ����� ���������� ����� OOS, OOS2 � �.�.
                ' ���������� ����������� ��������� ���, ����� ��� �� ��������� 12 ��������
                symbol.SetName "OOSTMP" & Right(CStr(currentSymbolId), 6)
                e3App.PutInfo 0, "  �������� ������������ � 'OOSTMP" & Right(CStr(currentSymbolId), 6) & "'."
            Else
                e3App.PutInfo 0, "  OOS ������ '" & symbolName & "' (ID: " & currentSymbolId & ") �� �������� �� �����. �� �� ����� ������������ ���� ��������."
            End If
        End If
    Next

    If OOSCountPlaced = 0 Then
        e3App.PutInfo 0, "�� ������� ����������� OOS �������� �� ����� ��� ���������� � ��������������."
        e3App.PutInfo 0, "=== ����� ������� ==="
        Set symbol = Nothing
        Set sheet = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    e3App.PutInfo 0, "������� " & OOSCountPlaced & " ����������� OOS ��������. ���������� �� ������� (�� ������� � �������)..."

    ' === ���������� OOSSymbolsToRename (Bubble Sort) ===
    ' (SymbolID, SheetID, SheetName, Column, Row, X, Y, OriginalName)
    ' Index:   0       1         2       3      4   5   6        7
    For i = 1 To OOSCountPlaced - 1
        For j = i + 1 To OOSCountPlaced
            Dim item1Array, item2Array
            item1Array = OOSSymbolsToRename(i)
            item2Array = OOSSymbolsToRename(j)

            ' Comparison logic for ASCENDING order: SheetName (numerically), then Column (string), then Row (string), then X (numeric), then Y (numeric)
            ' If item1 should come AFTER item2 in ASCENDING order, we swap.
            
            ' ��������� �� ����� ����� (��������)
            If CLng(item1Array(2)) > CLng(item2Array(2)) Then
                Call SwapArrayElements(OOSSymbolsToRename, i, j)
            ElseIf CLng(item1Array(2)) = CLng(item2Array(2)) Then
                ' ��������� �� ������� (���������, ��� ����� ��������)
                If StrComp(item1Array(3), item2Array(3), vbTextCompare) > 0 Then
                    Call SwapArrayElements(OOSSymbolsToRename, i, j)
                ElseIf StrComp(item1Array(3), item2Array(3), vbTextCompare) = 0 Then
                    ' ��������� �� ������ (���������, ��� ����� ��������)
                    If StrComp(item1Array(4), item2Array(4), vbTextCompare) > 0 Then
                        Call SwapArrayElements(OOSSymbolsToRename, i, j)
                    ElseIf StrComp(item1Array(4), item2Array(4), vbTextCompare) = 0 Then
                        ' ��������� �� X ������� (��������)
                        If item1Array(5) > item2Array(5) Then
                            Call SwapArrayElements(OOSSymbolsToRename, i, j)
                        ElseIf item1Array(5) = item2Array(5) Then
                            If item1Array(6) > item2Array(6) Then ' Compare Y position (ascending)
                                Call SwapArrayElements(OOSSymbolsToRename, i, j)
                            End If
                        End If
                    End If
                End If
            End If
        Next
    Next
    e3App.PutInfo 0, "���������� ���������. ���������� � �������������� �������������� OOS ��������."

    ' === ������������� �������������� OOS �������� � ��������������� ������� ===
    For s = 1 To OOSCountPlaced
        OOSCounter = OOSCounter + 1
        currentSymbolId = OOSSymbolsToRename(s)(0) ' Get Symbol ID from sorted array
        Dim originalSymbolName : originalSymbolName = OOSSymbolsToRename(s)(7) ' Get original name for logging
        
        symbol.SetId(currentSymbolId)
        
        Dim newSymbolName : newSymbolName = "OOS" & OOSCounter
        
        ' ��������������� ������
        symbol.SetName newSymbolName
        e3App.PutInfo 0, "  ������ '" & originalSymbolName & "' (ID: " & currentSymbolId & ") ������������ � '" & newSymbolName & "' (�� �����: " & OOSSymbolsToRename(s)(2) & " " & OOSSymbolsToRename(s)(3) & OOSSymbolsToRename(s)(4) & ")."
    Next

    e3App.PutInfo 0, "=== ���������� �������: ������������� " & OOSCounter & " OOS �������� ==="

    Set symbol = Nothing
    Set sheet = Nothing
    Set job = Nothing
    Set e3App = Nothing
End Sub

' Helper Sub: Swaps two elements in an array
Sub SwapArrayElements(arr, index1, index2)
    Dim temp
    temp = arr(index1)
    arr(index1) = arr(index2)
    arr(index2) = temp
End Sub

' === �������� ������ ===
Call RenameOOSSymbolsByLocation()
