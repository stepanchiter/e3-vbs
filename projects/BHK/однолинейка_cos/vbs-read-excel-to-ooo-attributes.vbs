Option Explicit

' === ������� ��������� === ������ ��������� OOS �� Excel
Sub WriteOOSAttributesFromExcel()
    Dim e3App, job, symbol
    Dim excelApp, excelWorkbook, excelSheet
    Dim filePath, i, rowNum, cellValue

    ' ���������� ��� �������� ���������
    Dim OOSTag, OOSType, OOSPras, OOSPnom, OOSInom
    Dim OOSDProizv3
    Dim OOSIras
    Dim OOSDProizv2 ' ���������� ��� �������� �� D_Proizv2
    Dim OOSDProizv1 ' ���������� ��� �������� �� D_Proizv1
    Dim OOSCos ' ���������� ��� �������� �� E_Cos

    Set e3App = CreateObject("CT.Application")
    Set job = e3App.CreateJobObject()
    Set symbol = job.CreateSymbolObject()

    e3App.PutInfo 0, "=== ����� �������: ������ ��������� OOS �� Excel ==="

    ' 1. ������ ���� � ����� XLSX
    filePath = InputBox("������� ������ ���� � ����� XLSX � �������:", "���� � ����� Excel", "C:\MyData\OOSAttributes.xlsx")

    If Trim(filePath) = "" Then
        e3App.PutInfo 0, "���� � ����� �� ��� ������. ������ �������."
        Exit Sub
    End If

    ' �������� ������������� �����
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        e3App.PutInfo 0, "������: ���� '" & filePath & "' �� ������. ��������� ����."
        Set fso = Nothing
        Exit Sub
    End If
    Set fso = Nothing

    ' 2. ������ Excel � �������� �����
    On Error Resume Next ' �������� ��������� ������ ��� �������� � Excel
    Set excelApp = GetObject("Excel.Application")
    If Err.Number <> 0 Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0 ' ��������� ��������� ������

    If excelApp Is Nothing Then
        e3App.PutInfo 0, "������: �� ������� ��������� ��� ������������ � Excel. ���������, ��� Excel ����������."
        Exit Sub
    End If

    excelApp.Visible = False ' �������� Excel ��� ������� ������
    excelApp.DisplayAlerts = False ' ��������� �������������� (��������, � ������ �������������)

    On Error Resume Next
    Set excelWorkbook = excelApp.Workbooks.Open(filePath)
    If Err.Number <> 0 Then
        e3App.PutInfo 0, "������: �� ������� ������� ���� Excel: '" & filePath & "'. ������: " & Err.Description
        excelApp.Quit
        Set excelApp = Nothing
        Exit Sub
    End If
    On Error GoTo 0

    Set excelSheet = excelWorkbook.Sheets(1) ' �������� � ������ ������

    ' 3. ��������� ���� �������� OOS �� ������� � �� ����������
    Dim allSymbolIds()
    Dim allSymbolCount
    ' job.GetSymbolIds ���������� ���������� ��������� � ��������� ������ allSymbolIds
    allSymbolCount = job.GetSymbolIds(allSymbolIds)

    If allSymbolCount = 0 Then
        e3App.PutInfo 0, "� ������� ��� �������� ��� �������. ������ ��������."
        excelWorkbook.Close False ' ������� ��� ����������
        excelApp.Quit
        Set excelSheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Exit Sub
    End If

    ' ���������� ������� ��� ���������� �������� OOS �������� � �� �������� ��������
    Dim OOSSymbolsMap
    Set OOSSymbolsMap = CreateObject("Scripting.Dictionary")
    Dim OOSNamesArray() ' ��� �������� �������� �������� OOS, ������� ����� �����������
    Dim OOSArrayCurrentSize : OOSArrayCurrentSize = 0 ' ������� ���������� ��������� � OOSNamesArray

    For i = LBound(allSymbolIds) To UBound(allSymbolIds) ' ���������� LBound � UBound ��� ����������
        symbol.SetId(allSymbolIds(i))
        Dim symName : symName = symbol.GetName()
        If LCase(Left(symName, 3)) = "OOS" Then
            ' ��������� �������� ������ �� ����� OOS (��������, �� "OOS12" �������� 12)
            Dim OOSNum : OOSNum = CLng(Mid(symName, 4))
            If Not OOSSymbolsMap.Exists(CStr(OOSNum)) Then
                OOSSymbolsMap.Add CStr(OOSNum), allSymbolIds(i)
                
                ' ����������� ������ ������� � ������������� ���
                OOSArrayCurrentSize = OOSArrayCurrentSize + 1
                ReDim Preserve OOSNamesArray(OOSArrayCurrentSize - 1) ' ��� 0-���������������� �������: N ��������� -> ������������ ������ N-1
                
                OOSNamesArray(OOSArrayCurrentSize - 1) = OOSNum ' ����������� �������� ���������� ��������
            Else
                e3App.PutInfo 0, "��������������: ��������� ������������� OOS ������ � ������� '" & OOSNum & "'. ����� ��������� ������ ������ ���������."
            End If
        End If
    Next

    If OOSSymbolsMap.Count = 0 Then
        e3App.PutInfo 0, "� ������� �� ������� �������� OOS ��� ������ ���������. ������ ��������."
        excelWorkbook.Close False
        excelApp.Quit
        Set excelSheet = Nothing
        Set excelWorkbook = Nothing
        Set excelApp = Nothing
        Set symbol = Nothing
        Set job = Nothing
        Set e3App = Nothing
        Set OOSSymbolsMap = Nothing
        Exit Sub
    End If

    ' ��������� �������� ������� OOS �������� �� �����������
    Call SortNumericArrayAsc(OOSNamesArray) ' ���������� ��������������� ������� ��� ����������

    e3App.PutInfo 0, "������� " & OOSSymbolsMap.Count & " �������� OOS. �������� ������ ���������..."

    ' 4. ���� �� ��������������� OOS �������� � ������ ���������
    For i = LBound(OOSNamesArray) To UBound(OOSNamesArray) ' ���������� LBound � UBound ��� �������� �� ���������������� �������
        Dim currentOOSNum : currentOOSNum = OOSNamesArray(i)
        Dim currentOOSId : currentOOSId = OOSSymbolsMap.Item(CStr(currentOOSNum))

        symbol.SetId(currentOOSId)
        Dim currentSymName : currentSymName = symbol.GetName()

        ' ������������� OOS_N -> ������ N+1
        ' ���� Excel ���������� �� ������ 1, � OOS_N � 1, �� ������ N+1 - ��� 2, 3 � �.�.
        ' ���� OOS_N ��� 0, �� ������ 0+1 = 1. ���������, ��� ��� ������������� ����� ��������� Excel.
        rowNum = currentOOSNum + 1 

        ' ������ �������� �� Excel
        On Error Resume Next ' �������� ��������� ������ ��� ������ �����
        OOSTag = Trim(CStr(excelSheet.Cells(rowNum, 13).Value))    ' ������� M --> �� E_TAG
        OOSType = Trim(CStr(excelSheet.Cells(rowNum, 14).Value))  ' ������� N --> �� E_TYPE
        OOSPras = Trim(CStr(excelSheet.Cells(rowNum, 6).Value)) ' ������� F --> �� E_Pras
        OOSPnom = Trim(CStr(excelSheet.Cells(rowNum, 5).Value))  ' ������� E --> �� E_Pnom
        OOSInom = Trim(CStr(excelSheet.Cells(rowNum, 8).Value)) ' ������� H --> �� E_Inom
        OOSDProizv3 = Trim(CStr(excelSheet.Cells(rowNum, 16).Value)) ' ������� P --> �� D_Proizv3
        OOSIras = Trim(CStr(excelSheet.Cells(rowNum, 9).Value))  ' ������� I --> �� E_Iras
        OOSDProizv2 = Trim(CStr(excelSheet.Cells(rowNum, 11).Value)) ' ������� K --> �� D_Proizv2
        OOSCos = Trim(CStr(excelSheet.Cells(rowNum, 4).Value))  ' ������� D --> �� E_Cos
        
        ' ������ "��������������� �������" �� "��" � ������
        If InStr(1, OOSDProizv2, "��������������� �������", vbTextCompare) > 0 Then
            OOSDProizv2 = Replace(OOSDProizv2, "��������������� �������", "��", 1, -1, vbTextCompare)
        End If
        
        ' ���� ������������ ��, �� � D_Proizv1 ���������� ���� ������������
        If InStr(1, OOSDProizv2, "��", vbTextCompare) > 0 Then
            OOSDProizv1 = "���� ������������ 24 VDC, 1 CO � �������� ���. RNC1CO024+SNB05-E-AR"
        Else
            OOSDProizv1 = "" ' ���� �� ��, �� ��������� ������
        End If
        
        If Err.Number <> 0 Then
            e3App.PutInfo 0, "��������������: ������ ��� ������ ������ ��� OOS ������� '" & currentSymName & "' (ID: " & currentOOSId & ") �� ������ " & rowNum & ". ��������� �������� ����� ���� �������. ������: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0

        ' ������ ��������� � ������ OOS
        e3App.PutInfo 0, "  ��������� �������: " & currentSymName & " (ID: " & currentOOSId & ") -> ������ Excel: " & rowNum

        If Len(OOSTag) > 0 Then
            symbol.SetAttributeValue "�� E_TAG", OOSTag
            e3App.PutInfo 0, "    �������� �� E_TAG: " & OOSTag
        Else
            e3App.PutInfo 0, "    �� E_TAG: <�����>"
        End If

        If Len(OOSType) > 0 Then
            symbol.SetAttributeValue "�� E_TYPE", OOSType
            e3App.PutInfo 0, "    �������� �� E_TYPE: " & OOSType
        Else
            e3App.PutInfo 0, "    �� E_TYPE: <�����>"
        End If

        If Len(OOSPras) > 0 Then
            symbol.SetAttributeValue "�� E_Pras", OOSPras
            e3App.PutInfo 0, "    �������� �� E_Pras: " & OOSPras
        Else
            e3App.PutInfo 0, "    �� E_Pras: <�����>"
        End If

        If Len(OOSPnom) > 0 Then
            symbol.SetAttributeValue "�� E_Pnom", OOSPnom
            e3App.PutInfo 0, "    �������� �� E_Pnom: " & OOSPnom
        Else
            e3App.PutInfo 0, "    �� E_Pnom: <�����>"
        End If

        If Len(OOSInom) > 0 Then
            symbol.SetAttributeValue "�� E_Inom", OOSInom
            e3App.PutInfo 0, "    �������� �� E_Inom: " & OOSInom
        Else
            e3App.PutInfo 0, "    �� E_Inom: <�����>"
        End If

        If Len(OOSDProizv3) > 0 Then
            symbol.SetAttributeValue "�� D_Proizv3", OOSDProizv3
            e3App.PutInfo 0, "    �������� �� D_Proizv3: " & OOSDProizv3
        Else
            e3App.PutInfo 0, "    �� D_Proizv3: <�����>"
        End If

        If Len(OOSIras) > 0 Then
            symbol.SetAttributeValue "�� E_Iras", OOSIras
            e3App.PutInfo 0, "    �������� �� E_Iras: " & OOSIras
        Else
            e3App.PutInfo 0, "    �� E_Iras: <�����>"
        End If
        
        If Len(OOSDProizv2) > 0 Then
            symbol.SetAttributeValue "�� D_Proizv2", OOSDProizv2
            e3App.PutInfo 0, "    �������� �� D_Proizv2: " & OOSDProizv2
        Else
            e3App.PutInfo 0, "    �� D_Proizv2: <�����>"
        End If
        
        If Len(OOSDProizv1) > 0 Then
            symbol.SetAttributeValue "�� D_Proizv1", OOSDProizv1
            e3App.PutInfo 0, "    �������� �� D_Proizv1: " & OOSDProizv1
        Else
            e3App.PutInfo 0, "    �� D_Proizv1: <�����>"
        End If
        
        If Len(OOSCos) > 0 Then
            symbol.SetAttributeValue "�� E_Cos", OOSCos
            e3App.PutInfo 0, "    �������� �� E_Cos: " & OOSCos
        Else
            e3App.PutInfo 0, "    �� E_Cos: <�����>"
        End If
    Next

    e3App.PutInfo 0, "=== ���������� �������: �������� ������� �������� ==="

    ' 5. ������� �������� Excel
    excelWorkbook.Close False ' ������� ��� ����������
    excelApp.Quit
    
    Set excelSheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
    Set symbol = Nothing
    Set job = Nothing
    Set e3App = Nothing
    Set OOSSymbolsMap = Nothing
End Sub

' === ��������������� ��������� === ���������� ��������� ������� �� �����������
Sub SortNumericArrayAsc(arr)
    Dim i, j, temp
    ' ���� ������ ���� ��� �������� ������ 1 �������, ���������� �� �����
    ' LBound(arr) � UBound(arr) ��������� ������������ ������ ����� ����������
    If UBound(arr) < LBound(arr) + 1 Then Exit Sub

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next
    Next
End Sub

' === �������� ������ ������� ===
Call WriteOOSAttributesFromExcel()