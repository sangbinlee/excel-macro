Sub PasteToVisibleCellsOnly()
    Dim rngSource As Range
    Dim rngTarget As Range
    Dim cell As Range
    Dim i As Long
    
    ' 붙여넣을 대상 영역 먼저 선택하도록 강제
    On Error Resume Next
    Set rngTarget = Application.InputBox("붙여넣을 영역(필터링된 상태)을 선택하세요", Type:=8)
    On Error GoTo 0
    
    ' 선택하지 않으면 종료
    If rngTarget Is Nothing Then
        MsgBox "붙여넣을 영역을 선택하지 않았습니다. 매크로를 종료합니다.", vbExclamation
        Exit Sub
    End If
    
    ' 복사할 범위 선택
    Set rngSource = Application.InputBox("복사할 범위를 선택하세요", Type:=8)
    
    ' 붙여넣기 실행 (보이는 셀만)
    i = 1
    For Each cell In rngTarget.SpecialCells(xlCellTypeVisible)
        cell.Value = rngSource.Cells(i, 1).Value
        i = i + 1
        If i > rngSource.Cells.Count Then Exit For
    Next cell
End Sub
