Sub SaveTemplateWithCellValue()

    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim wb As Workbook
    Dim selectedValue As String
    Dim rightValue As String
    Dim leftValue As String
    
    Dim templatePath As String
    Dim newFileName As String
    
    ' 1. get selected cell value
    If TypeName(Selection) <> "Range" Then
        MsgBox "셀을 선택하세요."
        Exit Sub
    End If
    
    selectedValue = Selection.Value
    
    '  the cell to the right of the selected cell (열 기준 +1)
    rightValue = Selection.Offset(0, 1).Value
    
    '  the cell to the left of the selected cell (열 기준 -1)
    leftValue = Selection.Offset(0, -1) '
    
    If selectedValue = "" Then
        MsgBox "The selected cell has no value."
        Exit Sub
    End If
    
    ' 2. Load template.xlsx  from the current folder
    templatePath = ThisWorkbook.Path & "\template.xlsx"
    If Dir(templatePath) = "" Then
        MsgBox "template.xlsx The file is not in the current folder."
        Exit Sub
    End If
    
    Set newWb = Workbooks.Open(templatePath)
        
    ' 3. Put a value into cell B2 on Sheet1
    newWb.Sheets(1).Range("B2").Value = Format(Date, "yyyy-mm-dd")
    ' 3. Put a value into cell D2 on Sheet1
    newWb.Sheets(1).Range("D2").Value = selectedValue
    ' 3. Put a value into cell E2 on Sheet1
    newWb.Sheets(1).Range("E2").Value = rightValue
    ' 3. Put a value into cell F2 on Sheet1
    newWb.Sheets(1).Range("F2").Value = leftValue
    
    ' 4. Put a value into cell B2 on Sheet2
    newWb.Sheets(2).Range("B2").Value = selectedValue
    
    ' 5. Put a value into cell C1 on Sheet3 + 시트명 변경
    newWb.Sheets(3).Range("C1").Value = selectedValue
    ' 5. Rename Sheet3
    newWb.Sheets(3).Name = selectedValue
        
    
    ' new file name set and save and open new file
    newFileName = ThisWorkbook.Path & "\make-app_" & selectedValue & "_" & leftValue & ".xlsx"
    
    Application.DisplayAlerts = False ' set no alert
    newWb.SaveAs Filename:=newFileName, FileFormat:=xlOpenXMLWorkbook
    newWb.Close
    Application.DisplayAlerts = True
    MsgBox "The file has been saved: " & newFileName
    
    Set wb = Workbooks.Open(newFileName) ' new file open
    wb.Activate ' new file activate
End Sub


