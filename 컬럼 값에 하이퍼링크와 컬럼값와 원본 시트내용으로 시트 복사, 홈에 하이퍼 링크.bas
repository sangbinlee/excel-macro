'  할일목록 만들기
Sub FillTodoList()
    Dim startValue As String
    Dim prefix As String
    Dim num As Integer
    Dim count As Integer
    Dim i As Integer
    Dim numPart As String
    Dim numLength As Integer
    
    ' A2 셀의 최초값 가져오기
    'startValue = Range("A2").Value
    
    ' 최초 문자열 입력받기 (예: todo-001)
    startValue = InputBox("최초 값을 입력하세요 (예: todo-001)", "시작값 입력")
    
        ' 입력값이 없으면 기본값 사용
    If Trim(startValue) = "" Then
        startValue = "todo-001"
    End If
    
    
     ' 숫자가 포함되어 있는지 확인
    If InStr(startValue, "-") > 0 Then
        ' "todos-001" 같은 경우: prefix와 숫자 분리
        prefix = Left(startValue, InStr(startValue, "-"))
        numPart = Mid(startValue, InStr(startValue, "-") + 1)
        num = CLng(numPart)
        numLength = Len(numPart)   ' 숫자 자리수 자동 감지
        
    Else
        ' "todos" 같은 경우: 자동으로 -001 붙여줌
        prefix = startValue & "-"
        num = 1
        numLength = 3              ' 기본 3자리
        'startValue = prefix & Format(num, "000")
        startValue = prefix & Format(num, String(numLength, "0"))
    End If
    
     
    ' 사용자에게 갯수 입력받기
    count = InputBox("몇 개를 생성하시겠습니까?", "갯수 입력", 10)
    
    ' A2부터 채우기
    'Range("A2").Value = startValue
    Range("A2").Value = prefix & Format(num, String(numLength, "0"))
    For i = 1 To count - 1
        'Range("A2").Offset(i, 0).Value = prefix & Format(num + i, "000")
        Range("A2").Offset(i, 0).Value = prefix & Format(num + i, String(numLength, "0"))
    Next i
End Sub

' 되돌리기 없으므로 매크로 함수로 해야함 , 할일목록 삭제
Sub ClearTodoList()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    'ws.Range("A1:A10").ClearContents
        
    ' 사용자에게 갯수 입력받기
    'count = InputBox("몇 개를 삭제하시겠습니까?", "갯수 입력", 10)
     
        
    ' A열에서 마지막 데이터 행 찾기
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' A2부터 마지막 행까지 지우기
    If lastRow >= 2 Then
        ws.Range("A2:A" & lastRow).ClearContents
    End If
    
End Sub

' 할일목록으로 시트 만들고 , 하이퍼링크 생성, 원본시트 복사 후 홈으로 가기 링크 만들기
Sub CreateSheetsAndHyperlinks()
    Dim wsIndex As Worksheet
    Dim cell As Range
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim sourceSheet As Worksheet
    
    ' index 시트 지정
    Set wsIndex = ThisWorkbook.Sheets("home")
    Set sourceSheet = ThisWorkbook.Sheets("원본")
    
    
    
    
    ' A열에서 마지막 데이터 행 찾기
    lastRow = wsIndex.Cells(wsIndex.Rows.count, 1).End(xlUp).Row
    
    
    ' A2 ~ A31 반복
    For Each cell In wsIndex.Range("A2:A" & lastRow)
        sheetName = Trim(cell.Value)
        
        If sheetName <> "" Then
            ' 시트 존재 여부 확인
            On Error Resume Next
            Set newSheet = ThisWorkbook.Sheets(sheetName)
            On Error GoTo 0
            
            ' 없으면 새 시트 생성
            If newSheet Is Nothing Then
                sourceSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
                Set newSheet = ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
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

'시트의 하이퍼링크와 생성된 시트가 모두 제거
Sub DeleteSheetsAndHyperlinks()
    Dim wsIndex As Worksheet
    Dim cell As Range
    Dim sheetName As String
    Dim ws As Worksheet
    
    ' index 시트 지정
    Set wsIndex = ThisWorkbook.Sheets("home")
    ' A열에서 마지막 데이터 행 찾기
    lastRow = wsIndex.Cells(wsIndex.Rows.count, 1).End(xlUp).Row
    
    ' 1. index 시트의 A2~A31 범위 하이퍼링크 삭제
    For Each cell In wsIndex.Range("A2:A" & lastRow)
        If cell.Hyperlinks.count > 0 Then
            cell.Hyperlinks.Delete
        End If
    Next cell
    
    ' 2. index 시트의 A2~A31에 적힌 이름의 시트 삭제
    Application.DisplayAlerts = False ' 삭제 확인 메시지 숨기기
    For Each cell In wsIndex.Range("A2:A" & lastRow)
        sheetName = Trim(cell.Value)
        If sheetName <> "" Then
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(sheetName)
            If Not ws Is Nothing Then
                ws.Delete
            End If
            On Error GoTo 0
            Set ws = Nothing
        End If
    Next cell
    Application.DisplayAlerts = True
End Sub

