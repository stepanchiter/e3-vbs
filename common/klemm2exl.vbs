'==============================================
' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜˜˜ XT, ˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜ ˜ Excel
' ˜˜˜. [˜˜˜˜˜˜˜ ˜˜˜˜]
'==============================================

' ˜˜˜˜˜˜˜˜ ˜˜˜
Set App = CreateObject("CT.Application")
Set Job = App.CreateJobObject
Set Dev = Job.CreateDeviceObject

' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜ ˜ Excel-˜˜˜˜˜ (˜˜˜˜˜˜˜˜ ˜˜ ˜˜˜ ˜˜˜˜˜˜˜˜ ˜˜˜˜)
'Dim excelFilePath
'excelFilePath = "\\Vt5\niokr\ZAKAZI2016\ZN_1108066\brk66976_241073206.xls"  ' ˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜ ˜˜˜˜˜!

' ˜˜˜˜˜˜˜˜˜ ˜˜˜ ˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜˜˜˜
Dim excelFilePath
excelFilePath = GetExcelFileName(Job)
If excelFilePath = "" Then
    WScript.Quit
End If

' ˜˜˜˜˜˜˜˜ ID ˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜
terminalCount = Job.GetTerminalIds(DevIds)

If terminalCount = 0 Then
    App.PutInfo 1, "˜˜˜˜˜˜˜˜˜ ˜˜ ˜˜˜˜˜˜˜!"
    WScript.Quit
End If

' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜
Dim xtTerminals
xtTerminals = ProcessTerminals(Job, Dev, DevIds, terminalCount)

' ˜˜˜˜˜˜˜˜˜, ˜˜˜ ˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜˜˜ ˜˜˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜˜˜
If IsArray(xtTerminals) Then
    ' ˜˜˜˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜ ˜˜˜˜˜˜˜˜
    Call ShowArrayBeforeExport(xtTerminals)
    
    ' ˜˜˜˜˜˜˜ ˜ Excel ˜ ˜˜˜˜˜˜˜˜˜
    Call ExportToExcel(xtTerminals, excelFilePath, "˜˜˜˜˜˜˜˜˜˜", "E2")
Else
    App.PutInfo 1, "˜˜˜ ˜˜˜˜˜˜ ˜˜˜ ˜˜˜˜˜˜˜˜ (XT-˜˜˜˜˜˜˜˜˜ ˜˜ ˜˜˜˜˜˜˜)"
End If

' ˜˜˜˜˜˜˜˜˜˜
Set Dev = Nothing
Set Job = Nothing
Set App = Nothing
WScript.Quit

'==============================================
' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜
' ˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ XT ˜˜˜˜˜˜˜˜˜˜ ˜˜˜ Nothing
'==============================================
Function ProcessTerminals(Job, Dev, DevIds, totalCount)
    ' 1. ˜˜˜˜ ˜˜˜˜˜˜
    Dim terminals(), xtTerminals()
    Dim i, j, xtCount
    
    ReDim terminals(totalCount - 1, 1)
    xtCount = 0
    
    For i = 0 To totalCount - 1
        Dev.SetId DevIds(i + 1)
        terminals(i, 0) = Dev.GetMasterPinName
        terminals(i, 1) = Dev.GetName
        
        If InStr(1, UCase(terminals(i, 1)), "XT", vbTextCompare) > 0 Then
            xtCount = xtCount + 1
        End If
    Next
    
    If xtCount = 0 Then
        ProcessTerminals = Null ' ˜˜˜˜˜˜˜˜˜˜ Null ˜˜˜˜˜˜ Nothing
        Exit Function
    End If
    
    ' 2. ˜˜˜˜˜˜˜˜˜˜ XT ˜˜˜˜˜˜˜˜˜˜
    ReDim xtTerminals(xtCount - 1, 1)
    j = 0
    
    For i = 0 To totalCount - 1
        If InStr(1, UCase(terminals(i, 1)), "XT", vbTextCompare) > 0 Then
            xtTerminals(j, 0) = terminals(i, 0)
            xtTerminals(j, 1) = terminals(i, 1)
            j = j + 1
        End If
    Next
    
    ' 3. ˜˜˜˜˜˜˜˜˜˜
    Dim temp1, temp2
    For i = UBound(xtTerminals, 1) - 1 To 0 Step -1
        For j = 0 To i
            If StrComp(xtTerminals(j, 1), xtTerminals(j + 1, 1), vbTextCompare) > 0 Then
                temp1 = xtTerminals(j, 0)
                temp2 = xtTerminals(j, 1)
                xtTerminals(j, 0) = xtTerminals(j + 1, 0)
                xtTerminals(j, 1) = xtTerminals(j + 1, 1)
                xtTerminals(j + 1, 0) = temp1
                xtTerminals(j + 1, 1) = temp2
            End If
        Next
    Next
    
    ProcessTerminals = xtTerminals
End Function


'==============================================
' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜˜˜˜˜ ˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜
'==============================================
Sub ShowArrayBeforeExport(dataArray)
    Dim i, outputStr
    outputStr = "˜˜˜˜˜˜ ˜˜˜ ˜˜˜˜˜˜˜˜ (" & UBound(dataArray, 1) + 1 & " ˜˜˜˜˜):" & vbCrLf
    outputStr = outputStr & "=========================" & vbCrLf
    outputStr = outputStr & "˜˜˜˜˜˜ | ˜˜˜˜˜ ˜˜˜˜˜˜ | ˜˜˜. ˜˜˜˜˜˜˜˜˜˜˜" & vbCrLf
    outputStr = outputStr & "-------------------------" & vbCrLf
    
    For i = 0 To UBound(dataArray, 1)
        outputStr = outputStr & Right("   " & i, 4) & " | " & _
                   Right("      " & dataArray(i, 0), 11) & " | " & _
                   dataArray(i, 1) & vbCrLf
    Next
    
    outputStr = outputStr & "========================="
    App.PutInfo 0, outputStr
End Sub

'==============================================
' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜ ˜ Excel (˜ ˜˜˜˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜˜)
'==============================================
Sub ExportToExcel(dataArray, filePath, sheetName, startCell)
    On Error Resume Next
    
    Dim ExcelApp, ExcelBook, ExcelSheet
    Dim i, colNum, rowNum, lastRow
    
    ' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜ ˜˜˜˜˜˜
    colNum = Asc(UCase(Mid(startCell, 1, 1))) - 64 ' A=1, B=2...
    rowNum = CInt(Mid(startCell, 2))
    
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.Visible = True ' ˜˜˜ ˜˜˜˜˜˜˜
    
    ' ˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜ ˜˜˜˜ ˜˜ ˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜
    App.PutInfo 0, "˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜ ˜˜˜˜: " & filePath
    Set ExcelBook = ExcelApp.Workbooks.Open(filePath)
    
    If Err.Number <> 0 Then
        App.PutInfo 1, "˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜ ˜˜˜˜˜ " & filePath & ": " & Err.Description
        ExcelApp.Quit
        Set ExcelApp = Nothing
        WScript.Quit
    End If
    On Error GoTo 0
    
    ' ˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜˜
    Set ExcelSheet = Nothing
    On Error Resume Next
    Set ExcelSheet = ExcelBook.Sheets(sheetName)
    If Err.Number <> 0 Then
        App.PutInfo 1, "˜˜˜˜ '" & sheetName & "' ˜˜ ˜˜˜˜˜˜ ˜ ˜˜˜˜˜ " & filePath
        ExcelBook.Close False
        ExcelApp.Quit
        Set ExcelApp = Nothing
        WScript.Quit
    End If
    On Error GoTo 0
    
    ' ˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜ E ˜˜˜˜˜˜˜ ˜ E2
    App.PutInfo 0, "˜˜˜˜˜˜ ˜˜˜˜˜˜˜ E ˜˜˜˜˜ ˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜..."
    With ExcelSheet
        ' ˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜ E
        lastRow = .Cells(.Rows.Count, 5).End(-4162).Row ' -4162 = xlUp
        
        ' ˜˜˜˜ ˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜˜ E2 - ˜˜˜˜˜˜˜
        If lastRow >= rowNum Then
            .Range(.Cells(rowNum, 5), .Cells(lastRow, 5)).ClearContents
        End If
    End With
    
    ' ˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜
    App.PutInfo 0, "˜˜˜˜˜˜˜˜˜ " & UBound(dataArray, 1) + 1 & " ˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜ E..."
    For i = 0 To UBound(dataArray, 1)
        ExcelSheet.Cells(rowNum + i, colNum).Value = dataArray(i, 0)
    Next
    
    ' ˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜ E ˜ ˜˜˜˜˜˜ ˜˜˜˜ (˜˜ E2 ˜˜ ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜)
    App.PutInfo 0, "˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜ E ˜ ˜˜˜˜˜˜ ˜˜˜˜..."
    With ExcelSheet
        lastRow = .Cells(.Rows.Count, 5).End(-4162).Row
        If lastRow < rowNum Then lastRow = rowNum ' ˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜, ˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ E2
        
        With .Range(.Cells(rowNum, 5), .Cells(lastRow, 5)).Interior
            .Color = 65535 ' ˜˜˜˜˜˜ ˜˜˜˜
            .Pattern = 1   ' ˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End With
    
    ' ˜˜˜˜˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜˜˜
    ExcelBook.Save
    ExcelBook.Close
    ExcelApp.Quit
    
    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelApp = Nothing
    
    App.PutInfo 0, "˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜: ˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜ E ˜˜˜˜˜˜˜˜ ˜ " & filePath
End Sub


'==============================================
' ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜ XLS ˜˜˜˜˜
' ˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜˜ ˜˜˜˜ ˜ ˜˜˜˜˜ ˜ ˜˜˜˜˜˜˜:
' <˜˜˜˜_˜˜˜˜˜˜˜>\<˜˜˜_˜˜˜˜˜˜˜>.xls
' ˜ ˜˜˜˜˜˜˜ "Sch2_" ˜˜ "brk" ˜ ˜˜˜˜˜˜˜˜˜ ".e3d"
'==============================================
Function GetExcelFileName(Job)
    Dim projectPath, projectName, excelFileName
    
    ' ˜˜˜˜˜˜˜˜ ˜˜˜˜ ˜˜˜˜˜˜˜
    projectPath = Job.GetPath()
    If Len("" & projectPath) = 0 Then
        App.PutInfo 1, "˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜ ˜˜˜˜ ˜˜˜˜˜˜˜"
        GetExcelFileName = ""
        Exit Function
    End If
    
    ' ˜˜˜˜˜˜˜˜ ˜˜˜ ˜˜˜˜˜˜˜
    projectName = Job.GetName()
    If Len("" & projectName) = 0 Then
        App.PutInfo 1, "˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜ ˜˜˜˜˜ ˜˜˜˜˜˜˜"
        GetExcelFileName = ""
        Exit Function
    End If
    
    ' ˜˜˜˜˜˜˜˜ "Sch2_" ˜˜ "brk" ˜ ˜˜˜˜˜ ˜˜˜˜˜˜˜
    projectName = Replace(projectName, "Sch2_", "brk", 1, -1, vbTextCompare)
    
    ' ˜˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜˜ .e3d ˜˜˜˜ ˜˜˜˜ (˜˜˜˜˜˜˜˜˜˜˜˜˜˜˜˜˜˜)
    projectName = Replace(projectName, ".e3d", "", 1, -1, vbTextCompare)
    projectName = Replace(projectName, ".E3D", "", 1, -1, vbTextCompare)
    
    ' ˜˜˜˜˜˜˜ ˜˜˜ ˜˜˜˜˜ Excel (˜˜˜ ˜˜˜˜˜˜ ˜˜˜˜˜˜˜˜˜)
    excelFileName = projectPath & "\" & projectName & ".xls"
    
    App.PutInfo 0, "˜˜˜˜˜˜˜˜˜˜˜ ˜˜˜˜ ˜ ˜˜˜˜˜ Excel: " & excelFileName
    GetExcelFileName = excelFileName
End Function