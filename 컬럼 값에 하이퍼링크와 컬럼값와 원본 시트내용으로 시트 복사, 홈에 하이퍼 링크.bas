Sub CreateSheetsAndHyperlinks()
    Dim wsIndex As Worksheet
    Dim cell As Range
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim sourceSheet As Worksheet
    
    ' index 시트 지정
    Set wsIndex = ThisWorkbook.Sheets("home")
    Set sourceSheet = ThisWorkbook.Sheets("원본")
    
    ' A2 ~ A31 반복
    For Each cell In wsIndex.Range("A2:A31")
        sheetName = Trim(cell.Value)
        
        If sheetName <> "" Then
            ' 시트 존재 여부 확인
            On Error Resume Next
            Set newSheet = ThisWorkbook.Sheets(sheetName)
            On Error GoTo 0
            
            ' 없으면 새 시트 생성
            If newSheet Is Nothing Then
                sourceSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                newSheet.Name = sheetName
            End If
            
            ' index 시트의 셀에 새 시트로 이동하는 하이퍼링크 추가
            wsIndex.Hyperlinks.Add Anchor:=cell, _
                Address:="", _
                SubAddress:="'" & sheetName & "'!A1", _
                TextToDisplay:=sheetName
            
            ' 새 시트의 B2 셀에 index 시트의 해당 셀로 돌아가는 하이퍼링크 추가
            newSheet.Hyperlinks.Add Anchor:=newSheet.Range("B2"), _
                Address:="", _
                SubAddress:="'" & wsIndex.Name & "'!" & cell.Address, _
                TextToDisplay:="home"
        End If
        
        Set newSheet = Nothing
    Next cell
End Sub

