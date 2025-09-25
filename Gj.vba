Sub AddVLookupFormulaRow11_Manual()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim i As Long
    
    Set wb = ActiveWorkbook
    
    For i = 3 To wb.Worksheets.Count
        Set ws = wb.Worksheets(i)
        lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
        
        For col = 2 To lastCol
            ' 단순하게 값만 입력 (수동으로 수식 완성)
            ws.Cells(11, col).Value = "=수식입력필요"
        Next col
    Next i
    
    MsgBox "위치 표시 완료! 수동으로 수식을 복사하세요.", vbInformation
End Sub
